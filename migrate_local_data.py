"""一次性迁移脚本:把桌面版(D:\\DevTools\\InventoryDesktopApp)本地 CSV 里的
真实数据——门店、分类、字段、物品、历史记录、管理员/员工账号——搬到新部署的
Google Apps Script 共享后端上。

只读本地数据,不会修改/删除本地 CSV 任何东西,可以放心重复尝试。

用法:
    1. 先按 google-apps-script.gs 文件头的步骤部署好新后端,拿到 /exec 网址。
    2. 直接运行 `python migrate_local_data.py`,按提示粘贴网址、给管理员和
       每个员工账号设置新密码(旧密码哈希算法跟新后端不一样,没法直接沿用,
       所以每个账号都需要一个新密码——迁移完了告诉对应的人就行)。

注意:历史记录的"消耗量/花费"是新后端按当前物品单价重新算出来的,不是照抄
本地 CSV 里当时算好的数字——这跟桌面版本身的行为一致(消耗/花费从来都是
"记录那一刻的单价"算出来直接存死,而不是能反推历史单价),只是搬家时用的是
物品"现在"的单价,这是数据本身就没法避免的极限,不是脚本的 bug。
"""
import getpass
import json
import os
import sys
import urllib.request

sys.path.insert(0, r"D:\DevTools\InventoryDesktopApp")
import inventory_core as core


def api_post(api_base_url, action, payload, token=None):
    body = dict(payload)
    body["action"] = action
    if token:
        body["token"] = token
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        api_base_url, data=data,
        headers={"Content-Type": "text/plain;charset=utf-8"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    if result.get("error"):
        raise RuntimeError(f"{action} 失败: {result['error']}")
    return result


def main():
    api_base_url = input("粘贴新后端的 /exec 网址(跟 config.js 里填的一样): ").strip()
    if not api_base_url:
        print("没有输入网址,已取消。")
        return

    local_stores = core.list_stores()
    local_users = core.list_users()
    admin_users = [u for u in local_users if u["role"] == "admin"]
    employee_users = [u for u in local_users if u["role"] == "employee"]

    print(f"本地发现: {len(local_stores)} 个门店 / {len(admin_users)} 个管理员账号 / {len(employee_users)} 个员工账号。")

    if not admin_users:
        print("本地没有管理员账号,没什么好迁移的。如果是全新开始,直接打开新网页版走"
              "\"设置管理员账号\"就行,不用跑这个脚本。")
        return

    admin = admin_users[0]
    admin_password = getpass.getpass(f"给管理员账号「{admin['username']}」设置新密码: ")
    try:
        reg = api_post(api_base_url, "registerAdmin", {
            "username": admin["username"], "password": admin_password, "displayName": admin["displayName"],
        })
        token = reg["token"]
        print(f"管理员账号「{reg['user']['username']}」已创建。")
    except RuntimeError as e:
        if "已经存在" not in str(e):
            raise
        print("新系统里已经有一个管理员账号了(可能之前在网页上手动注册过)。")
        existing_username = input("请输入那个已注册的管理员用户名: ").strip()
        existing_password = getpass.getpass("请输入对应的密码: ")
        login_result = api_post(api_base_url, "login", {
            "username": existing_username, "password": existing_password,
        })
        token = login_result["token"]
        print(f"已用现有管理员账号「{login_result['user']['username']}」登录,继续迁移。")

    store_id_map = {}
    for s in local_stores:
        created = api_post(api_base_url, "createStore", {"name": s["name"]}, token)["store"]
        store_id_map[s["id"]] = created["id"]
        print(f"门店「{s['name']}」已创建 -> {created['id']}")

    for s in local_stores:
        new_store_id = store_id_map[s["id"]]
        core.set_store(s["id"])
        categories = core.list_categories()
        for cat in categories:
            created_cat = api_post(api_base_url, "createCategory", {
                "storeId": new_store_id, "name": cat["name"],
            }, token)["category"]
            new_cat_id = created_cat["id"]

            settlement_currency = cat.get("settlementCurrency")
            if settlement_currency and settlement_currency != "CNY":
                api_post(api_base_url, "setSettlementCurrency", {
                    "categoryId": new_cat_id, "currency": settlement_currency,
                }, token)

            field_id_map = {}
            for f in core.list_fields(cat["id"]):
                created_field = api_post(api_base_url, "addField", {
                    "categoryId": new_cat_id, "name": f["name"], "fieldType": f["type"], "role": f["role"],
                }, token)["field"]
                field_id_map[f["id"]] = created_field["id"]

            local_items = core.list_items(cat["id"])
            item_id_map = {}
            item_prices = {}
            for it in local_items:
                created_item = api_post(api_base_url, "addItem", {
                    "categoryId": new_cat_id, "name": it["name"], "unit": it["unit"],
                    "price": it["price"], "currency": it["currency"],
                }, token)["item"]
                item_id_map[it["id"]] = created_item["id"]
                item_prices[it["id"]] = it["price"]

            log_df = core._read_log_df(cat["id"])
            log_df = log_df.sort_values("date")
            pushed = 0
            for _, row in log_df.iterrows():
                old_item_id = str(row["item_id"])
                if old_item_id not in item_id_map:
                    continue
                values = core._parse_values(row["values_json"])
                remapped_values = {field_id_map.get(k, k): v for k, v in values.items()}
                entry = {
                    "itemId": item_id_map[old_item_id],
                    "itemName": row["item_name"],
                    "values": remapped_values,
                    "price": item_prices.get(old_item_id, 0),
                }
                api_post(api_base_url, "saveLogEntries", {
                    "categoryId": new_cat_id, "date": row["date"], "entries": [entry],
                }, token)
                pushed += 1
            print(f"分类「{cat['name']}」迁移完成: {len(field_id_map)} 个字段, "
                  f"{len(item_id_map)} 个物品, {pushed} 条历史记录。")

    for emp in employee_users:
        new_store_id = store_id_map.get(emp["storeId"])
        if not new_store_id:
            print(f"跳过员工账号「{emp['username']}」:找不到它绑定的门店。")
            continue
        emp_password = getpass.getpass(f"给员工账号「{emp['username']}」设置新密码: ")
        api_post(api_base_url, "createEmployee", {
            "username": emp["username"], "password": emp_password,
            "displayName": emp["displayName"], "storeId": new_store_id,
        }, token)
        print(f"员工账号「{emp['username']}」已创建,记得把新密码告诉她。")

    print("迁移完成。本地 CSV 原样保留,没有被修改。")


if __name__ == "__main__":
    main()
