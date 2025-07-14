from sheet_parser import SheetParser

def main():
    parser = SheetParser()

    # 解析Excel文件
    try:
        html = parser.parse_file('test1.xlsx', 'test1.html', title='示例表格')
        print("Excel文件解析完成")
    except Exception as e:
        print(f"解析Excel文件时出错: {e}")

if __name__ == "__main__":
    main()
