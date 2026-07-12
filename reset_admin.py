"""临时诊断脚本:如果登录一直提示密码不对,怀疑是隐藏密码输入框打字看不见导致
打错——这个脚本改用明文输入(你打的字会直接显示在屏幕上,不隐藏),确保设置的
密码和你等下登录用的密码是同一个。

用法: python reset_admin.py
"""
import json
import urllib.request


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
    api_base_url = input("粘贴后端网址(config.js 里那个): ").strip()
    username = input("管理员用户名(比如 949696481): ").strip()
    password = input("管理员密码(会直接显示在屏幕上,确保没打错): ").strip()
    display_name = input("显示名字(比如 Xin): ").strip() or username

    reg = api_post(api_base_url, "registerAdmin", {
        "username": username, "password": password, "displayName": display_name,
    })
    print(f"注册成功: {reg['user']}")
    print(f"记住这个密码,等下登录网页版就用这个: {password}")

    print("\n马上验证一下能不能用这个密码登录...")
    login_result = api_post(api_base_url, "login", {"username": username, "password": password})
    print(f"验证成功,登录正常: {login_result['user']}")


if __name__ == "__main__":
    main()
