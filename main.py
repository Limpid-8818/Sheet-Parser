from sheet_parser import SheetParser


def main():
    parser = SheetParser()

    try:
        html = parser.parse_file('examples/basic_example.xlsx', 'examples/example.html', title='示例表格')
        print("Excel文件解析完成")
    except Exception as e:
        print(f"解析Excel文件时出错: {e}")

    try:
        html = parser.parse_file('examples/example_csv.csv', 'examples/example_csv.html', title='CSV示例表格')
        print("CSV文件解析完成")
    except Exception as e:
        print(f"解析CSV文件时出错: {e}")


if __name__ == "__main__":
    main()
