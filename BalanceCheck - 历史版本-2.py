import sys
import tkinter as tk
from tkinter import scrolledtext
from tkinter import messagebox
import pandas as pd
import os

class DataProcessingApp:
    def __init__(self, master):
        self.master = master
        master.title("内部往来核对程序")

        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(master, width=80, height=10)
        self.log_text.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        # 数据处理按钮
        self.process_button = tk.Button(master, text="执行数据处理", command=self.process_data, width=20)
        self.process_button.grid(row=1, column=0, padx=25, pady=5)

        # 行项目匹配按钮
        self.match_button = tk.Button(master, text="凭证行项目匹配", command=self.execute_matching, width=20)
        self.match_button.grid(row=1, column=1, padx=25, pady=5)

        # 规划求解按钮
        self.solve_button = tk.Button(master, text="执行规划求解算法", command=self.solve_problem, width=20)
        self.solve_button.grid(row=1, column=2, columnspan=2, padx=25, pady=5)


        # 初始化合并后的数据
        self.merged_data_data = None

    def process_data(self):
        messagebox.showinfo("提示", "若数据量较大，可能需要执行一段时间，程序运行过程中，不要点击其他按钮")
        self.log("开始读取、拼接、转换数据...")
        try:
            # 获取当前脚本所在目录
            current_directory = os.path.dirname(os.path.abspath(__file__))

            # 读取四个Excel文件
            customers_data = pd.read_excel(os.path.join(current_directory, 'customers_data.xlsx'))
            customers_tot = pd.read_excel(os.path.join(current_directory, 'customers_tot.xlsx'))
            suppliers_data = pd.read_excel(os.path.join(current_directory, 'suppliers_data.xlsx'))
            suppliers_tot = pd.read_excel(os.path.join(current_directory, 'suppliers_tot.xlsx'))

            # 将suppliers_tot的数据从第二行开始拼接在customers_tot的数据下方
            merged_data_tot = pd.concat([customers_tot, suppliers_tot.iloc[1:]], ignore_index=True)
            # 将suppliers_data的数据从第二行开始拼接在customers_data的数据下方
            merged_data_data = pd.concat([customers_data, suppliers_data.iloc[1:]], ignore_index=True)

            # 合并"客户"列和"供应商"列的数据
            merged_data_tot['客商编码'] = merged_data_tot['客户'].fillna(merged_data_tot['供应商'])
            merged_data_data['客商编码'] = merged_data_data['客户'].fillna(merged_data_data['供应商'])

            # 转换为字符串类型
            merged_data_tot['客商编码'] = merged_data_tot['客商编码'].fillna("NaN").astype(str)
            merged_data_data['客商编码'] = merged_data_data['客商编码'].fillna("NaN").astype(str)

            # 去除最后两位字符
            merged_data_tot['客商编码'] = merged_data_tot['客商编码'].apply(lambda x: x[:-2] if len(x) > 2 else x)
            merged_data_data['客商编码'] = merged_data_data['客商编码'].apply(lambda x: x[:-2] if len(x) > 2 else x)

            # 创建df存储转换规则
            conversion_rules = {
                '8000': '9000',
                '3001': '9001',
                '8002': '9002',
                '8003': '9003',
                '8005': '9005',
                '1000': '1000',
                '1100': '1100',
                '1200': '1200',
                '1201': '1201',
                '1301': '1301',
                '1500': '1500',
                '1501': '1501',
                '1502': '1502',
                '1601': '1601',
                '1602': '1602',
                '1700': '1700',
                '1701': '1701',
                '1800': '1800',
                '1801': '1801',
                '1900': '1900',
                '2000': '2000',
                '2999': '2999',
                '10002892': '2100',
                '10003727': '2100',
                '2300': '2300',
                '2301': '2301',
                '10004026': '2400',
                '10004120': '2401',
                '10004117': '2402',
                '2403': '2403',
                '10004618': '2404',
                '10004460': '2405',
                '10004511': '2406',
                '2406': '2406',
                '10004482': '2407',
                '10004754': '2408',
                '2408': '2408',
                '2409': '2409',
                '2500': '2500',
                '4019': '3100',
                '4032': '3200',
                '3201': '3201',
                '3202': '3202',
                '4038': '3300',
                '4040': '3301',
                '3302': '3302',
                '4022': '3400',
                '4045': '3500',
                '4016': '3600',
                '4049': '3701',
                '4014': '3800',
                '4031': '3900',
                '8200322': '9000',
                '8401193': '9001',
                '8402045': '9002',
                '8402765': '9003',
                '8404009': '9005',
                '8100008': '1000',
                '8400171': '1100',
                '8200323': '1200',
                '8402658': '1201',
                '8400717': '1301',
                '8400802': '1500',
                '8402085': '1501',
                '8402786': '1502',
                '8400677': '1601',
                '8200872': '1602',
                '8400384': '1700',
                '8404215': '1701',
                '8400164': '1800',
                '8100031': '1801',
                '8401453': '1900',
                '8400191': '2000',
                '8405391': '2999',
                '8403628': '2100',
                '8400675': '2300',
                '8402230': '2301',
                '8403717': '2400',
                '8300004': '2401',
                '8100034': '2402',
                '8404925': '2403',
                '8405247': '2404',
                '8405240': '2405',
                '8404891': '2406',
                '8405195': '2406',
                '8405243': '2407',
                '8405445': '2408',
                '8405998': '2409',
                '8404236': '2500',
                '8100006': '3100',
                '8200399': '3200',
                '8100030': '3201',
                '8404502': '3202',
                '8100011': '3300',
                '8100013': '3301',
                '8404218': '3302',
                '8100005': '3400',
                '8100020': '3500',
                '8100003': '3600',
                '8200852': '3701',
                '8100001': '3800',
                '8200392': '3900',

            }

            # 定义转换函数
            def convert_code(code):
                return conversion_rules.get(code, code)  # 如果找不到匹配规则，则返回原始代码

            # 应用转换函数到"客商编码"列
            merged_data_tot['内部账套编码转换'] = merged_data_tot['客商编码'].apply(convert_code)
            merged_data_data['内部账套编码转换'] = merged_data_data['客商编码'].apply(convert_code)
            
            # 创建一个新列，用于凭证行项目对账
            merged_data_data['匹配批次'] = ''
            # 将公司代码列转换为字符串格式
            merged_data_data['公司代码'] = merged_data_data['公司代码'].astype(str)

            # 创建透视表
            pivot_table = pd.pivot_table(merged_data_tot, index='公司代码', columns='内部账套编码转换', values='公司代码货币价值', aggfunc='sum', fill_value=0)

            # 设置合并后的数据
            self.merged_data_data = merged_data_data
            
            #开始处理合并数据
            # 提取并合并公司代码列和内部账套编码转换列
            company_codes = pd.concat([merged_data_tot['公司代码'].astype(str), merged_data_tot['内部账套编码转换'].astype(str)], ignore_index=True)

            # 去除重复值
            company_codes_unique = company_codes.drop_duplicates().reset_index(drop=True).sort_values()

            # 创建一个新的DataFrame，初始化为NaN
            company_codes_unique_df = pd.DataFrame(index=range(len(company_codes_unique) + 1), columns=range(len(company_codes_unique) + 1))

            # 在第一列和第一行上分别填入去重后的数据
            company_codes_unique_df.iloc[1:, 0] = company_codes_unique.values
            company_codes_unique_df.iloc[0, 1:] = company_codes_unique.values
            # 输出成功信息
            self.log("数据处理完成，执行创建excel文件...")
            # 创建一个新的Excel writer
            with pd.ExcelWriter(os.path.join(current_directory, '内部往来数据核对.xlsx')) as writer:

                
                # 将拼接后的数据保存到新的工作表中
                merged_data_tot.to_excel(writer, index=False, sheet_name='科目维度汇总底表')
                merged_data_data.to_excel(writer, index=False, sheet_name='凭证维度明细底表')
                
                # 创建透视表并保存到新的工作表中
                pivot_table.to_excel(writer, sheet_name='全账套内部往来汇总')
                
                # 创建新工作表，存储透视表行列转置后的数据合计
                company_codes_unique_df.to_excel(writer, sheet_name='当期内部往来', index=False,  header=False)
            
            self.log("文件创建完毕，命名为'内部往来数据核对.xlsx'")
            messagebox.showinfo("提示", "执行完毕！")
        except Exception as e:
            messagebox.showerror("错误", f"处理数据时发生错误：{str(e)}")
            self.log("处理数据时发生错误！")

    def execute_matching(self):
        messagebox.showinfo("提示", "若数据量较大，可能需要执行一段时间，程序运行过程中，不要点击其他按钮，程序可以后台运行，出现未响应也不需要做任何操作")
        self.log("开始执行行项目匹配...")
        try:
            # 确保数据已经处理
            if self.merged_data_data is None:
                raise ValueError("未处理数据，请先执行数据处理！")

            # 创建一个空的匹配批次列表
            batch_list = []
            

            # 初始化流水号
            serial_number = 0
            batch_prefix = 'P'

            # 初始化成功和失败的计数器
            success_count = 0
            failure_count = 0

            # 循环遍历数据框中的每一行
            for index, row in self.merged_data_data.iterrows():
                # 如果已经存在匹配批次，则跳过
                if row['匹配批次']:
                    continue

                # 提取当前行的公司代码、内部账套编码转换和公司代码货币价值
                company_code = row['公司代码']
                account_code = row['内部账套编码转换']
                currency_value = row['公司代码货币价值']

                # 查找匹配条件
                match_condition = self.merged_data_data[(self.merged_data_data['公司代码'] == account_code) & 
                                                   (self.merged_data_data['内部账套编码转换'] == company_code) & 
                                                   (self.merged_data_data['公司代码货币价值'] == -currency_value)]

                # 如果找到匹配项且匹配批次为空，则将匹配批次设置为当前流水号
                if not match_condition.empty and not match_condition.iloc[0]['匹配批次']:
                    match_index = match_condition.index[0]
                    match_batch = batch_prefix + str(serial_number)
                    self.log(f"匹配成功：行{index}和行{match_index}，匹配批次为{match_batch}")
                    batch_list.append(match_batch)
                    self.merged_data_data.at[index, '匹配批次'] = match_batch
                    self.merged_data_data.at[match_index, '匹配批次'] = match_batch
                    serial_number += 1
                    success_count += 1
                else:
                    self.log(f"未找到匹配项或匹配批次已存在：行{index}")
                    failure_count += 1

            # 成功的总数和失败的总数
            self.log(f"成功匹配总数：{success_count}")
            self.log(f"失败匹配总数：{failure_count}")
            self.log(f"在凭证维度明细底表中的最后一列，可以看到凭证行项目的匹配批次。")
            # 输出成功信息
            messagebox.showinfo("提示", "执行完毕！")
        except Exception as e:
            messagebox.showerror("错误", f"执行行项目匹配时发生错误：{str(e)}")
            self.log("执行行项目匹配时发生错误！")

    def solve_problem(self):
        messagebox.showinfo("提示", "规划求解功能暂未启用！")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.master.update_idletasks()

def main():
    root = tk.Tk()
    app = DataProcessingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
