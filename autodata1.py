import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd
import configparser


class ExcelHandler(FileSystemEventHandler):
    def __init__(self, config_file):
        self.config = configparser.ConfigParser()
        self.config.read(config_file)
        self.input_folder = self.config.get('Paths', 'InputFolder')
        self.output_folder = self.config.get('Paths', 'OutputFolder')

    def on_created(self, event):
        if event.is_directory:
            return
        # 去除_设备
        if event.src_path.endswith("_设备.xls") or event.src_path.endswith("_设备.xlsx"):
            return
        if event.src_path.endswith('.xls') or event.src_path.endswith('.xlsx'):
            time.sleep(1)  # 等待1秒钟
            excel_path = event.src_path
            self.process_excel(excel_path)

    def process_excel(self, excel_path):
        try:
            # 读取Excel数据并进行操作
            df = pd.read_excel(excel_path, skiprows=1)  # 替换为您的Excel文件路径

            # 统计数据
            total_rows = len(df)  # 总行数
            total_columns = len(df.columns)  # 总列数
            sum_column1 = df['序号'].sum()  # 统计 Column1 列的和
            average_column2 = df['成功率'].mean()  # 统计 Column2 列的平均值
            max_column3 = df['成功率'].max()  # 统计 成功率 列的最大值
            min_column4 = df['成功率'].min()  # 统计 成功率 列的最小值
            max_column5 = df['平均响应时长'].max()  # 统计 平均响应时长的最大值
            max_column6 = df['平均响应时长'].min()  # 统计 平均响应时长的最小值

            # 打印统计结果、
            print(f"总行数：{total_rows}")
            print(f"总列数：{total_columns}")
            print(f"Column1 列的和：{sum_column1}")
            print(f"Column2 列的平均值：{average_column2}")
            print(f"Column3 列的最大值：{max_column3}")
            print(f"Column4 列的最小值：{min_column4}")

            # 将处理结果保存到另一个路径下
            output_file = os.path.join(self.output_folder, 'processed_data.txt')
            with open(output_file, 'w') as f:
                f.write(f"总行数：{total_rows}\n")
                f.write(f"总列数：{total_columns}\n")
                f.write(f"成功率的平均值：{average_column2}\n")
                f.write(f"成功率列的最大值：{max_column3}\n")
                f.write(f"成功率的最小值：{min_column4}\n")
                f.write(f"平均响应时长的最大值：{max_column5}\n")
                f.write(f"平均响应时长的最小值：{max_column6}\n")

            os.remove(excel_path)  # 删除原始Excel文件

        except Exception as e:
            print(f"An error occurred: {e}")


def main():
    config_file = 'config.ini'  # 配置文件路径
    excel_processor = ExcelHandler(config_file)

    observer = Observer()
    observer.schedule(excel_processor, excel_processor.input_folder, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
            print("Waiting for new Excel files...")
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    while True:
        main()
