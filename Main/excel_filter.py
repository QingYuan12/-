import pandas as pd
import os
import re
import sys  # 添加这行
import tkinter as tk
from tkinter import ttk, messagebox


class ExcelDataFilter:
    def __init__(self, file_path):
        """
        初始化，读取Excel文件
        """
        self.file_path = file_path
        self.data = {}
        self.load_excel_data()

    def load_excel_data(self):
        """
        读取Excel文件中的所有工作表数据
        """
        try:
            # 读取所有工作表
            xls = pd.ExcelFile(self.file_path)

            # 读取每个工作表
            for sheet_name in xls.sheet_names:
                # 读取工作表，不将第一行作为表头
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                # 使用map替换applymap
                df = df.map(lambda x: str(x) if pd.notna(x) else '')
                self.data[sheet_name] = df

            print(f"成功加载Excel文件: {self.file_path}")
            return True

        except Exception as e:
            print(f"读取Excel文件失败: {e}")
            return False

    def find_probability(self, df, layer, initial_node, target_node):
        """
        查找概率数据 - 扩大查找范围版本
        """
        # 查找层数行
        layer_row = -1
        for i in range(len(df)):
            if str(df.iloc[i, 1]).strip() == layer:
                layer_row = i
                break

        if layer_row == -1:
            return None

        # 从层数行开始查找所有初始节点行
        for i in range(layer_row, len(df)):
            # 检查是否到达下一个层数
            if i > layer_row and str(df.iloc[i, 1]).strip() in ['一层', '二~四层', '五层及以上', '树洞']:
                break

            # 查找初始节点标记行
            if str(df.iloc[i, 2]).strip() == "初始节点":
                # 检查这一行的所有初始节点
                for col in range(3, df.shape[1]):
                    init_cell = str(df.iloc[i, col]).strip()

                    if not init_cell:
                        continue

                    # 检查是否匹配初始节点
                    if '/' in init_cell:
                        nodes = [n.strip() for n in init_cell.split('/')]
                        if initial_node in nodes:
                            matched_init_row = i
                            matched_init_col = col
                            break
                    elif init_cell == initial_node:
                        matched_init_row = i
                        matched_init_col = col
                        break
                else:
                    continue  # 如果没有在这个初始节点行找到匹配的，继续下一个

                # 找到目标节点标记行
                target_marker_row = -1
                for j in range(matched_init_row + 1, min(matched_init_row + 10, len(df))):
                    if str(df.iloc[j, 2]).strip() == "目标节点":
                        target_marker_row = j
                        break

                if target_marker_row == -1:
                    continue

                # 目标节点值在目标节点标记行
                target_value_row = target_marker_row

                search_range = 8

                # 先检查同一列
                if matched_init_col < df.shape[1]:
                    target_cell = str(df.iloc[target_value_row, matched_init_col]).strip()

                    if target_cell == target_node:
                        # 数据行在下一行
                        data_row = target_value_row + 1
                        if data_row < len(df):
                            data_value = str(df.iloc[data_row, matched_init_col]).strip()
                            if data_value:
                                return data_value

                # 优先向右查找，从近到远
                for offset in range(1, search_range + 1):
                    check_col = matched_init_col + offset
                    if check_col >= df.shape[1]:
                        break

                    target_cell = str(df.iloc[target_value_row, check_col]).strip()

                    if target_cell == target_node:
                        # 数据行在下一行
                        data_row = target_value_row + 1
                        if data_row < len(df):
                            data_value = str(df.iloc[data_row, check_col]).strip()
                            if data_value:
                                return data_value

                # 然后向左查找，从近到远
                for offset in range(1, search_range + 1):
                    check_col = matched_init_col - offset
                    if check_col < 3:  # 至少是D列
                        break

                    target_cell = str(df.iloc[target_value_row, check_col]).strip()

                    if target_cell == target_node:
                        # 数据行在下一行
                        data_row = target_value_row + 1
                        if data_row < len(df):
                            data_value = str(df.iloc[data_row, check_col]).strip()
                            if data_value:
                                return data_value

                # 如果还没找到，尝试在整个行中查找（作为最后手段）
                for check_col in range(3, df.shape[1]):
                    target_cell = str(df.iloc[target_value_row, check_col]).strip()

                    if target_cell == target_node:
                        # 数据行在下一行
                        data_row = target_value_row + 1
                        if data_row < len(df):
                            data_value = str(df.iloc[data_row, check_col]).strip()
                            if data_value:
                                return data_value

        return None

    def calculate_formula(self, formula, df):
        """
        计算Excel公式
        """
        try:
            # 提取公式中的单元格引用
            cell_refs = re.findall(r'\$[A-Z]+\$\d+', formula)

            if not cell_refs:
                # 尝试非绝对引用
                cell_refs = re.findall(r'[A-Z]+\d+', formula)

            if not cell_refs:
                return formula

            # 获取单元格数值
            cell_values = {}
            for ref in cell_refs:
                # 解析引用
                if ref.startswith('$'):
                    # 绝对引用，如$B$2
                    col_letter = ref[1:].split('$')[0]
                    row_num = int(ref.split('$')[2])
                else:
                    # 相对引用，如B2
                    col_letter = re.match(r'[A-Z]+', ref).group()
                    row_num = int(re.search(r'\d+', ref).group())

                # 转换为行索引和列索引
                col_idx = ord(col_letter) - ord('A')
                row_idx = row_num - 1  # Excel行号从1开始

                if 0 <= row_idx < len(df) and 0 <= col_idx < df.shape[1]:
                    cell_val = str(df.iloc[row_idx, col_idx]).strip()
                    if cell_val and cell_val.replace('.', '').replace('-', '').isdigit():
                        cell_values[ref] = float(cell_val)
                    else:
                        cell_values[ref] = 0
                else:
                    cell_values[ref] = 0

            # 替换公式中的引用
            eval_formula = formula[1:]  # 去掉等号
            for ref, val in cell_values.items():
                eval_formula = eval_formula.replace(ref, str(val))

            # 安全地计算表达式
            try:
                result = eval(eval_formula, {"__builtins__": {}}, {})
                return result
            except Exception as e:
                print(f"公式计算错误: {e}")
                return formula

        except Exception as e:
            print(f"公式解析错误: {e}")
            return formula

    def format_as_percentage(self, value):
        """
        将数值格式化为百分数
        """
        try:
            if isinstance(value, (int, float)):
                # 如果是数值，转换为百分数
                if value <= 1:
                    # 如果小于等于1，认为是小数，乘以100
                    percentage = value * 100
                else:
                    # 如果大于1，直接作为百分数（如100表示100%）
                    percentage = value

                # 格式化为两位小数的百分数
                return f"{percentage:.2f}%"
            elif isinstance(value, str):
                # 如果是字符串，尝试转换为数值
                try:
                    num_value = float(value)
                    if num_value <= 1:
                        percentage = num_value * 100
                    else:
                        percentage = num_value
                    return f"{percentage:.2f}%"
                except:
                    # 如果无法转换，返回原始字符串
                    return value
            else:
                return str(value)
        except Exception as e:
            print(f"格式化百分数错误: {e}")
            return str(value)

    def get_probability(self, meeting_type, layer, initial_node, target_node):
        """
        获取指定条件下的概率
        """
        if meeting_type not in self.data:
            return None

        df = self.data[meeting_type]

        # 查找数据
        data_value = self.find_probability(df, layer, initial_node, target_node)

        if data_value is None or data_value == '':
            return 0

        # 处理数据值
        if data_value.startswith('='):
            # 尝试解析公式
            calculated_value = self.calculate_formula(data_value, df)

            if isinstance(calculated_value, (int, float)):
                # 如果是小数形式，直接返回
                if calculated_value <= 1:
                    return calculated_value
                else:
                    # 如果是百分数形式，转换为小数
                    return calculated_value / 100
            else:
                # 无法解析公式，尝试从字符串中提取数值
                try:
                    num_value = float(calculated_value)
                    if num_value <= 1:
                        return num_value
                    else:
                        return num_value / 100
                except:
                    return 0
        else:
            # 直接返回数值
            try:
                num_value = float(data_value)
                if num_value <= 1:
                    return num_value
                else:
                    return num_value / 100
            except:
                return 0


class ProbabilityCalculatorGUI:
    def __init__(self, excel_file):
        """
        初始化GUI界面
        """
        self.excel_file = excel_file
        self.filter_app = None

        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("节点刷新概率计算器")
        self.root.geometry("650x750")

        # 设置样式
        self.setup_styles()

        # 加载Excel数据
        self.load_excel_data()

        # 创建界面
        self.create_widgets()

    def setup_styles(self):
        """
        设置界面样式
        """
        style = ttk.Style()
        # 设置字体大小
        style.configure('Title.TLabel', font=('微软雅黑', 18, 'bold'))
        style.configure('Section.TLabel', font=('微软雅黑', 14, 'bold'))
        style.configure('Result.TLabel', font=('微软雅黑', 16, 'bold'), foreground='blue')

        # 为Radiobutton和Checkbutton创建自定义样式
        style.configure('Large.TRadiobutton', font=('微软雅黑', 12))
        style.configure('Large.TCheckbutton', font=('微软雅黑', 12))

    def load_excel_data(self):
        """
        加载Excel数据
        """
        try:
            if not os.path.exists(self.excel_file):
                messagebox.showerror("错误", f"找不到Excel文件 '{self.excel_file}'")
                self.root.destroy()
                return

            self.filter_app = ExcelDataFilter(self.excel_file)
            if not self.filter_app.data:
                messagebox.showerror("错误", "Excel文件加载失败")
                self.root.destroy()
                return

        except Exception as e:
            messagebox.showerror("错误", f"加载Excel文件时出错: {e}")
            self.root.destroy()
            return

    def create_widgets(self):
        """
        创建界面控件
        """
        # 标题
        title_label = ttk.Label(self.root, text="节点刷新概率计算器", style='Title.TLabel')
        title_label.pack(pady=20)

        # 创建框架容器
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 相遇类型选择
        meeting_frame = ttk.LabelFrame(main_frame, text="相遇类型", padding=10)
        meeting_frame.pack(fill=tk.X, pady=(0, 15))

        self.meeting_var = tk.StringVar(value="死仇")
        meetings = ["死仇", "美愿", "涂鸦", "无相遇"]
        for i, meeting in enumerate(meetings):
            rb = ttk.Radiobutton(meeting_frame, text=meeting, value=meeting,
                                 variable=self.meeting_var, style='Large.TRadiobutton')
            rb.grid(row=0, column=i, padx=15, pady=5)

        # 层数选择
        layer_frame = ttk.LabelFrame(main_frame, text="层数", padding=10)
        layer_frame.pack(fill=tk.X, pady=(0, 15))

        self.layer_var = tk.StringVar(value="树洞")
        layers = ["树洞", "一层", "二~四层", "五层及以上"]
        for i, layer in enumerate(layers):
            rb = ttk.Radiobutton(layer_frame, text=layer, value=layer,
                                 variable=self.layer_var, style='Large.TRadiobutton')
            rb.grid(row=0, column=i, padx=15, pady=5)

        # 初始节点选择
        initial_frame = ttk.LabelFrame(main_frame, text="初始节点", padding=10)
        initial_frame.pack(fill=tk.X, pady=(0, 15))

        self.initial_vars = {}
        initial_nodes = ["紧急", "作战", "不期", "安全", "先行", "失与得", "得偿", "商店"]
        for i, node in enumerate(initial_nodes):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(initial_frame, text=node, variable=var,
                                 style='Large.TCheckbutton')
            cb.grid(row=i // 4, column=i % 4, padx=15, pady=5, sticky="w")
            self.initial_vars[node] = var

        # 目标节点选择（允许多选）
        target_frame = ttk.LabelFrame(main_frame, text="目标节点（可多选）", padding=10)
        target_frame.pack(fill=tk.X, pady=(0, 15))

        self.target_vars = {}
        target_nodes = ["紧急", "作战", "不期", "安全", "先行", "失与得", "得偿"]
        for i, node in enumerate(target_nodes):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(target_frame, text=node, variable=var,
                                 style='Large.TCheckbutton')
            cb.grid(row=i // 4, column=i % 4, padx=15, pady=5, sticky="w")
            self.target_vars[node] = var

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        # 计算按钮
        self.calc_button = ttk.Button(button_frame, text="计算概率",
                                      command=self.calculate_probability, width=15)
        self.calc_button.pack(side=tk.LEFT, padx=10)

        # 清空按钮
        clear_button = ttk.Button(button_frame, text="清空选择",
                                  command=self.clear_selection, width=15)
        clear_button.pack(side=tk.LEFT, padx=10)

        # 结果框架
        result_frame = ttk.LabelFrame(main_frame, text="计算结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 结果文本区域
        self.result_text = tk.Text(result_frame, height=8, width=60,
                                   font=('微软雅黑', 14), bg='#f0f0f0')
        self.result_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var,
                               relief=tk.SUNKEN, anchor=tk.W, font=('微软雅黑', 10))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def get_selected_nodes(self, var_dict):
        """
        获取选中的节点列表
        """
        selected = []
        for node, var in var_dict.items():
            if var.get():
                selected.append(node)
        return selected

    def calculate_probability(self):
        """
        计算概率
        """
        # 获取选中的初始节点
        initial_selected = self.get_selected_nodes(self.initial_vars)
        if len(initial_selected) != 1:
            messagebox.showwarning("警告", "请选择且只能选择一个初始节点")
            return

        # 获取选中的目标节点
        target_selected = self.get_selected_nodes(self.target_vars)
        if len(target_selected) == 0:
            messagebox.showwarning("警告", "请至少选择一个目标节点")
            return

        # 获取选中的参数
        meeting_type = self.meeting_var.get()
        layer = self.layer_var.get()
        initial_node = initial_selected[0]

        # 更新状态
        self.status_var.set("正在计算...")
        self.calc_button.config(state='disabled')
        self.root.update()

        try:
            total_probability = 0
            detailed_results = []

            # 计算每个选中的目标节点的概率
            for target_node in target_selected:
                # 计算概率
                prob_value = self.filter_app.get_probability(meeting_type, layer, initial_node, target_node)

                if prob_value is None:
                    detailed_results.append(f"{target_node}: 0%")
                else:
                    # 格式化为百分数
                    formatted_value = self.filter_app.format_as_percentage(prob_value)
                    total_probability += prob_value
                    detailed_results.append(f"{target_node}: {formatted_value}")

            # 格式化总概率
            total_percentage = self.filter_app.format_as_percentage(total_probability)

            # 显示结果
            self.result_text.delete(1.0, tk.END)

            if len(target_selected) == 1:
                # 单目标节点，只显示结果
                self.result_text.insert(tk.END, total_percentage)
            else:
                # 多目标节点，显示详细结果和总和
                for result in detailed_results:
                    self.result_text.insert(tk.END, result + "\n")

                # 添加分隔线
                separator = "-" * 30 + "\n"
                self.result_text.insert(tk.END, separator)

                # 显示总和
                self.result_text.insert(tk.END, f"总和: {total_percentage}")

            self.status_var.set("计算完成")

        except Exception as e:
            messagebox.showerror("错误", f"计算过程中出现错误: {str(e)}")
            self.status_var.set("计算错误")

        finally:
            self.calc_button.config(state='normal')

    def clear_selection(self):
        """
        清空所有选择
        """
        # 重置初始节点选择
        for var in self.initial_vars.values():
            var.set(False)

        # 重置目标节点选择
        for var in self.target_vars.values():
            var.set(False)

        # 清空结果区域
        self.result_text.delete(1.0, tk.END)

        # 重置状态
        self.status_var.set("已清空选择")

    def run(self):
        """
        运行GUI应用
        """
        self.root.mainloop()


def main():
    """
    主函数
    """
    # 设置Excel文件路径
    excel_file = "target.xlsx"

    # 如果是打包后的exe，需要获取正确的路径
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        base_path = sys._MEIPASS
        excel_file = os.path.join(base_path, "target.xlsx")
    else:
        # 如果是直接运行Python脚本
        excel_file = "target.xlsx"

    # 创建并运行GUI应用
    app = ProbabilityCalculatorGUI(excel_file)
    app.run()


if __name__ == "__main__":
    main()