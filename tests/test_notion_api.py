import pytest
from notion_api import fetch_and_process_orders

# Notion APIからのレスポンスを模したダミーデータ
@pytest.fixture
def mock_notion_raw_data():
    return {
        "orders": [
            # Supplier A: 2 items
            {"supplier_name": "仕入先A", "db_part_number": "PART-001", "page_id": "page1"},
            {"supplier_name": "仕入先A", "db_part_number": "PART-002", "page_id": "page2"},
            # Supplier B: 1 item
            {"supplier_name": "仕入先B", "db_part_number": "PART-003", "page_id": "page3"},
            # Supplier is None: 1 item
            {"supplier_name": None, "db_part_number": "PART-004", "page_id": "page4"},
        ],
        "unlinked_count": 1
    }

def test_fetch_and_process_orders(monkeypatch, mock_notion_raw_data):
    """
    fetch_and_process_orders関数が正しくデータをグループ化できるかテストする
    """
    # notion_api.get_order_data_from_notionが、実際のAPI通信の代わりにダミーデータを返すように「すり替え」
    monkeypatch.setattr("notion_api.get_order_data_from_notion", lambda department_names=None: mock_notion_raw_data)

    # テスト対象の関数を実行
    result = fetch_and_process_orders(department_names=["営業部"])

    # --- 結果を検証 ---
    # 全体の注文数は4件であること
    assert len(result["all_orders"]) == 4

    # 未リンクの件数が1件であること
    assert result["unlinked_count"] == 1

    # グループ化された仕入先は2社であること
    assert len(result["orders_by_supplier"]) == 2

    # 仕入先Aの注文は2件であること
    assert len(result["orders_by_supplier"]["仕入先A"]) == 2

    # 仕入先Bの注文は1件であること
    assert len(result["orders_by_supplier"]["仕入先B"]) == 1

    # 仕入先がNoneのデータはグループに含まれないこと
    assert None not in result["orders_by_supplier"]
