from chart_creater import create_chart
import matplotlib.pyplot as plt
import argparse

def main():
    parser = argparse.ArgumentParser(description='製作圖表')
    parser.add_argument('--excel', type=str, default= 'data.xlsx',  help='使用Excel檔案來生成圖表')
    parser.add_argument('--title', type=str, default= '圖表', help=' 圖表的標題')
    parser.add_argument('--x', type=str,default="X axis", help='X軸的標籤')
    parser.add_argument('--y', type=str,default='Y axis', help='Y軸的標籤')
    parser.add_argument('--type', type=str, default='line', help='圖表的類型')
    args = parser.parse_args()
    create_chart(args.excel, args.title, args.x, args.y, args.type)
    plt.show()

if __name__ == '__main__':
    main()