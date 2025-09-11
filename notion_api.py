import time
from datetime import datetime
from notion_client import Client
import config

def get_order_data_from_notion():
    """Notionから「要発注」ステータスの注文データを取得する"""
    if not config.NOTION_API_TOKEN or not config.PAGE_ID_CONTAINING_DB:
        print("エラー: NotionのAPIトークンまたはデータベースIDが.envファイルに設定されていません。")
        return None
        
    notion = Client(auth=config.NOTION_API_TOKEN)
    order_list = []
    try:
        print(f"DBコンテナ ({config.PAGE_ID_CONTAINING_DB}) を検索中...")
        children = notion.blocks.children.list(block_id=config.PAGE_ID_CONTAINING_DB)
        real_database_id = next((b.get("id") for b in children.get("results", []) if b.get("type") == "child_database"), None)
        
        if not real_database_id:
            print(f"エラー: DBコンテナ内に子データベースが見つかりません。")
            return []
        print(f"DB発見: {real_database_id}")
        
        all_results = []
        next_cursor = None
        while True:
            query_res = notion.databases.query(database_id=real_database_id, start_cursor=next_cursor)
            all_results.extend(query_res.get("results", []))
            if not query_res.get("has_more"):
                break
            next_cursor = query_res.get("next_cursor")

        print(f"全 {len(all_results)} 件のデータをフィルタリング中...")
        for page in all_results:
            props = page.get("properties", {})
            # "要発注" ステータスをチェック
            if "要発注" not in props.get("注文ステータス", {}).get("formula", {}).get("string", ""):
                continue
            
            # 仕入先リレーションをチェック
            supplier_relation = props.get("DB_仕入先マスター", {}).get("relation", [])
            if not supplier_relation:
                continue
            supplier_page_id = supplier_relation[0].get("id")

            try:
                # APIレート制限を考慮
                time.sleep(0.35)
                supplier_page = notion.pages.retrieve(page_id=supplier_page_id)
                supplier_props = supplier_page.get("properties", {})
                
                # 必要なプロパティを安全に取得
                order_list.append({
                    "page_id": page["id"],
                    "maker_name": props.get("メーカー名", {}).get("rich_text", [{}])[0].get("plain_text", ""),
                    "db_part_number": props.get("DB品番", {}).get("rich_text", [{}])[0].get("plain_text", ""),
                    "quantity": props.get("数量", {}).get("number", 0),
                    "supplier_name": supplier_props.get("購入先", {}).get("title", [{}])[0].get("plain_text", ""),
                    "sales_contact": supplier_props.get("営業担当", {}).get("rich_text", [{}])[0].get("plain_text", ""),
                    "email": supplier_props.get("メール", {}).get("email", ""),
                    "email_cc": supplier_props.get("メール CC:", {}).get("email", ""),
                })
            except Exception as e:
                print(f"仕入先情報の取得中にエラーが発生しました (Page ID: {supplier_page_id}): {e}")
                
        print(f"-> フィルタリング完了。{len(order_list)} 件の要発注データが見つかりました。")
    except Exception as e:
        print(f"Notionデータベースの処理中にエラーが発生しました: {e}")
        
    return order_list

def update_notion_pages(page_ids):
    """指定されたNotionページのリストの「発注日」を今日の日付に更新する"""
    print(f"{len(page_ids)}件のNotionページの「発注日」を更新中...")
    notion = Client(auth=config.NOTION_API_TOKEN)
    today = datetime.now().strftime("%Y-%m-%d")
    
    for i, page_id in enumerate(page_ids):
        try:
            notion.pages.update(
                page_id=page_id,
                properties={"発注日": {"date": {"start": today}}}
            )
            print(f"({i+1}/{len(page_ids)}) {page_id} の更新完了")
            # APIレート制限を考慮
            time.sleep(0.35)
        except Exception as e:
            print(f"エラー: {page_id} の更新に失敗しました. {e}")
            
    print("Notionページの更新が完了しました。")
