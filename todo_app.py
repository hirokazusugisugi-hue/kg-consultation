import json
import os

TODO_FILE = "todos.json"

def load_todos():
    """JSONファイルからTodoリストを読み込む"""
    if os.path.exists(TODO_FILE):
        with open(TODO_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_todos(todos):
    """TodoリストをJSONファイルに保存"""
    with open(TODO_FILE, 'w', encoding='utf-8') as f:
        json.dump(todos, f, ensure_ascii=False, indent=2)

def show_todos(todos):
    """Todoリストを表示"""
    if not todos:
        print("\nTodoリストは空です")
        return

    print("\n" + "=" * 40)
    print("Todoリスト")
    print("=" * 40)
    for i, todo in enumerate(todos, 1):
        status = "✓" if todo['done'] else " "
        print(f"{i}. [{status}] {todo['task']}")
    print("=" * 40)

def add_todo(todos):
    """新しいタスクを追加"""
    task = input("新しいタスクを入力: ").strip()
    if task:
        todos.append({'task': task, 'done': False})
        save_todos(todos)
        print(f"追加しました: {task}")
    else:
        print("タスクが空です")

def complete_todo(todos):
    """タスクを完了にする"""
    if not todos:
        print("Todoリストは空です")
        return

    show_todos(todos)
    try:
        num = int(input("完了にするタスク番号を入力: "))
        if 1 <= num <= len(todos):
            todos[num - 1]['done'] = True
            save_todos(todos)
            print(f"完了しました: {todos[num - 1]['task']}")
        else:
            print("無効な番号です")
    except ValueError:
        print("数値を入力してください")

def delete_todo(todos):
    """タスクを削除"""
    if not todos:
        print("Todoリストは空です")
        return

    show_todos(todos)
    try:
        num = int(input("削除するタスク番号を入力: "))
        if 1 <= num <= len(todos):
            removed = todos.pop(num - 1)
            save_todos(todos)
            print(f"削除しました: {removed['task']}")
        else:
            print("無効な番号です")
    except ValueError:
        print("数値を入力してください")

def main():
    todos = load_todos()

    print("=" * 40)
    print("簡単なTodoリストアプリ")
    print("=" * 40)

    while True:
        print("\n操作を選択してください:")
        print("1. Todoリストを表示")
        print("2. タスクを追加")
        print("3. タスクを完了にする")
        print("4. タスクを削除")
        print("5. 終了")

        choice = input("\n選択 (1-5): ").strip()

        if choice == '1':
            show_todos(todos)
        elif choice == '2':
            add_todo(todos)
        elif choice == '3':
            complete_todo(todos)
        elif choice == '4':
            delete_todo(todos)
        elif choice == '5':
            print("アプリを終了します")
            break
        else:
            print("1-5の数字を入力してください")

if __name__ == "__main__":
    main()
