def calculator():
    print("簡単な電卓プログラム")
    print("使用できる演算子: +, -, *, /")
    print("終了するには 'q' を入力してください")
    print("-" * 40)

    while True:
        # 最初の数値を入力
        num1_input = input("\n最初の数値を入力 (qで終了): ")
        if num1_input.lower() == 'q':
            print("電卓を終了します")
            break

        try:
            num1 = float(num1_input)
        except ValueError:
            print("エラー: 有効な数値を入力してください")
            continue

        # 演算子を入力
        operator = input("演算子を入力 (+, -, *, /): ")
        if operator not in ['+', '-', '*', '/']:
            print("エラー: 有効な演算子を入力してください")
            continue

        # 2番目の数値を入力
        num2_input = input("2番目の数値を入力: ")
        try:
            num2 = float(num2_input)
        except ValueError:
            print("エラー: 有効な数値を入力してください")
            continue

        # 計算を実行
        if operator == '+':
            result = num1 + num2
            print(f"\n結果: {num1} + {num2} = {result}")
        elif operator == '-':
            result = num1 - num2
            print(f"\n結果: {num1} - {num2} = {result}")
        elif operator == '*':
            result = num1 * num2
            print(f"\n結果: {num1} × {num2} = {result}")
        elif operator == '/':
            if num2 == 0:
                print("エラー: ゼロで割ることはできません")
            else:
                result = num1 / num2
                print(f"\n結果: {num1} ÷ {num2} = {result}")

if __name__ == "__main__":
    calculator()
