import time
from datetime import datetime
from notion_client import Client
import config
import concurrent.futures

# --- 安全なデータ取得のためのヘルパー関数群 ---

def _get_safe_text(prop_list):
    """rich_textまたはtitleのリストから安全にテキストを取得する"""
    if prop_list and isinstance(prop_list, list) and len(prop_list) > 0:
        return prop_list[0].get('plain_text', '')
    return ''

def _get_safe_email(prop):
    """emailプロパティから安全に値を取得する"""
    return prop.get('email') if prop else ''

def _get_safe_number(prop):
    """numberプロパティから安全に値を取得する"""
    return prop.get('number', 0) if prop else 0


def _get_all_pages_from_db(notion_client, database_id, filter_params=None):
    """
    指定されたデータベースからすべてのページを取得する（ページネーション対応）。
    オプションでフィルターを適用可能。
    """
    all_results = []
    next_cursor = None
    while True:
        query_args = {
            "database_id": database_id,
            "start_cursor": next_cursor
        }
        if filter_params:
            query_args["filter"] = filter_params

        query_res = notion_client.databases.query(**query_args)
        all_results.extend(query_res.get("results", []))
        if not query_res.get("has_more"):
            break
        next_cursor = query_res.get("next_cursor")
    return all_results

def get_order_data_from_notion(department_names=None):
    """
    Notionから「要発注」ステータスの注文データを効率的に取得する。
    部署名によるフィルタリング機能を追加。
    """
    if not all([config.NOTION_API_TOKEN, config.PAGE_ID_CONTAINING_DB, config.NOTION_SUPPLIER_DATABASE_ID]):
        return None

    notion = Client(auth=config.NOTION_API_TOKEN)
    order_list = []
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            # 1. 仕入先DB取得と注文DB取得を並列で実行
            future_suppliers = executor.submit(_get_all_pages_from_db, notion, config.NOTION_SUPPLIER_DATABASE_ID)
            
            # --- フィルター条件の構築 ---
            # ベースとなるフィルター（総合発注判定プロパティを利用）
            base_filter = {
                "property": "総合発注判定",
                "formula": {"string": {"contains": "要発注"}}
            }

            # 部署名フィルターが指定されている場合の処理
            if department_names:
                department_filters = [
                    {"property": "部署名", "multi_select": {"contains": name}}
                    for name in department_names
                ]
                # 部署名フィルターが1つならそのまま、複数ならorで連結
                if len(department_filters) == 1:
                    department_filter_condition = department_filters[0]
                else:
                    department_filter_condition = {"or": department_filters}
                
                # 既存のフィルターとANDで結合
                final_filter = {
                    "and": [
                        base_filter,
                        department_filter_condition
                    ]
                }
            else:
                # 部署名が指定されていなければ、既存のフィルターのみ使用
                final_filter = base_filter

            # 注文DB取得をサブミット
            future_orders = executor.submit(_get_all_pages_from_db, notion, config.PAGE_ID_CONTAINING_DB, filter_params=final_filter)

            # 両方の処理が終わるのを待って結果を取得
            all_suppliers = future_suppliers.result()
            order_pages = future_orders.result()

        # 取得したデータを後処理
        suppliers_map = {page['id']: page['properties'] for page in all_suppliers}
        total_orders = len(order_pages)
        
        if total_orders == 0:
            return []

        # 3. 取得した注文データと仕入先データを結合
        for i, page in enumerate(order_pages):
            props = page.get("properties", {})
            
            supplier_relation = props.get("DB_仕入先リスト", {}).get("relation", [])
            if not supplier_relation:
                continue
            
            supplier_page_id = supplier_relation[0].get("id")
            supplier_props = suppliers_map.get(supplier_page_id)

            if not supplier_props:
                continue

            # ヘルパー関数を使って安全にデータを抽出
            order_list.append({
                "page_id": page["id"],
                "maker_name": _get_safe_text(props.get("メーカー名", {}).get("rich_text")),
                "db_part_number": _get_safe_text(props.get("DB品番", {}).get("rich_text")),
                "quantity": _get_safe_number(props.get("数量")),
                "supplier_name": _get_safe_text(supplier_props.get("購入先", {}).get("title")),
                "sales_contact": _get_safe_text(supplier_props.get("営業担当", {}).get("rich_text")),
                "email": _get_safe_email(supplier_props.get("メール")),
                "email_cc": _get_safe_email(supplier_props.get("メール CC:")),
            })
        
    except Exception as e:
        pass
        
    return order_list

def update_notion_pages(page_ids):
    """指定されたNotionページのリストの「発注日」を今日の日付に更新する"""
    if not page_ids:
        return
        
    notion = Client(auth=config.NOTION_API_TOKEN)
    today = datetime.now().strftime("%Y-%m-%d")
    
    for i, page_id in enumerate(page_ids):
        try:
            notion.pages.update(
                page_id=page_id,
                properties={"発注日": {"date": {"start": today}}}
            )
            time.sleep(0.35)
        except Exception as e:
            pass # app_gui.pyでエラーハンドリング
