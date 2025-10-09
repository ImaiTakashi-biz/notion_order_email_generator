import concurrent.futures
import threading
import time
from collections import defaultdict
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

import notion_client
from notion_client import Client

import config

_NOTION_LOCK = threading.Lock()
_NOTION_CLIENT: Optional[Client] = None
_NOTION_TOKEN: Optional[str] = None


def _get_notion_client() -> Client:
    """
    Notionクライアントをキャッシュし、トークン変更時のみ再生成する。
    """
    global _NOTION_CLIENT, _NOTION_TOKEN
    token = config.NOTION_API_TOKEN
    if _NOTION_CLIENT is None or token != _NOTION_TOKEN:
        _NOTION_CLIENT = notion_client.Client(auth=token)
        _NOTION_TOKEN = token
    return _NOTION_CLIENT


def _get_safe_text(prop_list: List[Dict[str, Any]]) -> str:
    """rich_text/title プロパティからテキストを安全に抽出する。"""
    if not prop_list or not isinstance(prop_list, list):
        return ""
    return "".join(item.get("plain_text", "") for item in prop_list)


def _get_safe_email(prop: Optional[Dict[str, Any]]) -> str:
    """email プロパティからアドレスを安全に抽出する。"""
    return (prop or {}).get("email", "")


def _get_safe_number(prop: Optional[Dict[str, Any]]) -> Union[int, float]:
    """number プロパティから値を安全に抽出する。"""
    return (prop or {}).get("number", 0)


def _get_all_pages_from_db(
    client: Client,
    database_id: str,
    filter_params: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    指定したデータベースから全ページを取得する。簡易リトライ付き。
    """
    all_results: List[Dict[str, Any]] = []
    next_cursor: Optional[str] = None

    while True:
        query_args: Dict[str, Any] = {"database_id": database_id, "start_cursor": next_cursor}
        if filter_params:
            query_args["filter"] = filter_params

        query_res: Optional[Dict[str, Any]] = None
        for attempt in range(3):
            try:
                with _NOTION_LOCK:
                    query_res = client.databases.query(**query_args)
                break
            except Exception:
                if attempt == 2:
                    return all_results
                time.sleep(config.AppConstants.NOTION_API_DELAY * (attempt + 1))

        if not query_res:
            return all_results

        all_results.extend(query_res.get("results", []))
        if not query_res.get("has_more"):
            break
        next_cursor = query_res.get("next_cursor")

    return all_results


def get_order_data_from_notion(department_names: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    発注対象データを Notion から取得する。
    """
    if not all(
        [config.NOTION_API_TOKEN, config.PAGE_ID_CONTAINING_DB, config.NOTION_SUPPLIER_DATABASE_ID]
    ):
        return {"orders": [], "unlinked_count": 0}

    client = _get_notion_client()
    order_list: List[Dict[str, Any]] = []
    unlinked_count = 0

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            future_suppliers = executor.submit(
                _get_all_pages_from_db, client, config.NOTION_SUPPLIER_DATABASE_ID
            )

            base_filter: Dict[str, Any] = {
                "property": "発注判定",
                "formula": {"string": {"contains": "要発注"}},
            }

            if department_names:
                department_filters = [
                    {"property": "部署名", "multi_select": {"contains": name}} for name in department_names
                ]
                department_condition = (
                    department_filters[0] if len(department_filters) == 1 else {"or": department_filters}
                )
                final_filter: Dict[str, Any] = {"and": [base_filter, department_condition]}
            else:
                final_filter = base_filter

            future_orders = executor.submit(
                _get_all_pages_from_db, client, config.PAGE_ID_CONTAINING_DB, filter_params=final_filter
            )

            all_suppliers = future_suppliers.result()
            order_pages = future_orders.result()

        suppliers_map = {page["id"]: page.get("properties", {}) for page in all_suppliers}
        if not order_pages:
            return {"orders": [], "unlinked_count": 0}

        for page in order_pages:
            props = page.get("properties", {})

            supplier_relation = props.get("DB_仕入先リスト", {}).get("relation", [])
            if not supplier_relation:
                unlinked_count += 1
                continue

            supplier_page_id = supplier_relation[0].get("id")
            supplier_props = suppliers_map.get(supplier_page_id)
            if not supplier_props:
                unlinked_count += 1
                continue

            department_entries = props.get("部署名", {}).get("multi_select", [])
            department_names_for_order = [
                entry.get("name", "").strip()
                for entry in department_entries
                if isinstance(entry, dict) and entry.get("name")
            ]

            maker = _get_safe_text(props.get("メーカー名", {}).get("rich_text")).strip()
            part_number = _get_safe_text(props.get("DB品番", {}).get("rich_text")).strip()
            quantity = int(_get_safe_number(props.get("数量")) or 0)
            remarks = _get_safe_text(props.get("備考", {}).get("rich_text")).strip()
            supplier_name = _get_safe_text(supplier_props.get("購入先", {}).get("title")).strip()
            sales_contact = _get_safe_text(supplier_props.get("営業担当", {}).get("rich_text")).strip()
            email_to = (_get_safe_email(supplier_props.get("メール")) or "").strip()
            email_cc = (_get_safe_email(supplier_props.get("メール CC:")) or "").strip()

            order_list.append(
                {
                    "page_id": page["id"],
                    "maker_name": maker,
                    "db_part_number": part_number,
                    "quantity": quantity,
                    "supplier_name": supplier_name,
                    "sales_contact": sales_contact,
                    "email": email_to,
                    "email_cc": email_cc,
                    "remarks": remarks,
                    "departments": department_names_for_order,
                }
            )

    except Exception:
        return {"orders": [], "unlinked_count": 0}

    return {"orders": order_list, "unlinked_count": unlinked_count}


def update_notion_pages(page_ids: List[str]) -> None:
    """
    対象ページの「発注日」を当日日付で更新する。
    """
    client = _get_notion_client()
    today = datetime.now().strftime("%Y-%m-%d")

    for page_id in page_ids:
        try:
            with _NOTION_LOCK:
                client.pages.update(page_id=page_id, properties={"発注日": {"date": {"start": today}}})
            time.sleep(config.AppConstants.NOTION_API_DELAY)
        except Exception:
            continue


def fetch_and_process_orders(department_names: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Notionから取得したデータを仕入先単位でグルーピングして返す。
    """
    raw_data = get_order_data_from_notion(department_names)
    orders = raw_data.get("orders", [])
    unlinked_count = raw_data.get("unlinked_count", 0)

    grouped_orders: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for order in orders:
        supplier_name = order.get("supplier_name")
        if supplier_name:
            grouped_orders[supplier_name].append(order)

    return {"orders_by_supplier": dict(grouped_orders), "all_orders": orders, "unlinked_count": unlinked_count}
