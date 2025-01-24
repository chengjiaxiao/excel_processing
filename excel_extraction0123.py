import pandas as pd
import os
import panel as pn
from io import BytesIO
pn.extension()

class ExcelMerger:
    def __init__(self):
        self.filepaths = None
        self.available_sheets = []
        self.selected_sheets = []
        
        # 创建UI组件
        self.file_input = pn.widgets.FileInput(accept='.xlsx,.xls,.xlsm', multiple=True)
        self.sheet_selector = pn.widgets.CheckBoxGroup(name='选择要合并的工作表', options=[])
        self.select_all_button = pn.widgets.Button(name='全选/取消全选', button_type='default')
        self.start_range_input = pn.widgets.TextInput(name='起始单元格 (例如: A1)', value='', placeholder='留空则读取全部内容')
        self.end_range_input = pn.widgets.TextInput(name='结束单元格 (例如: D10)', value='', placeholder='留空则等于起始单元格')
        self.use_header_checkbox = pn.widgets.Checkbox(name='使用第一行作为标题', value=False)
        self.merge_button = pn.widgets.Button(name='合并选中的工作表', button_type='primary')
        
        # 添加下载按钮
        self.download_button = pn.widgets.FileDownload(button_type='success', visible=False)
        
        # 绑定事件
        self.file_input.param.watch(self._update_sheets, 'value')
        self.select_all_button.on_click(self._toggle_all_sheets)
        self.merge_button.on_click(self._merge_sheets)
        
        # 创建界面布局
        self.layout = pn.Column(
            # pn.pane.Markdown('# Excel工作表合并工具'),
            pn.pane.Markdown('## 1. 请选择Excel文件'),
            self.file_input,
            pn.pane.Markdown('## 2. 选择要合并的工作表'),
            self.select_all_button,
            self.sheet_selector,
            pn.pane.Markdown('## 3. 指定单元格范围（可选）'),
            pn.Row(self.start_range_input, self.end_range_input),
            self.use_header_checkbox,
            pn.pane.Markdown('## 4. 开始合并'),
            pn.Row(self.merge_button, self.download_button)
        )

    def split_filename_tt(self, filename):
        parts = filename.split('-')
        return parts[1] if len(parts) > 1 else None

    def _update_sheets(self, event):
        """当文件被上传时更新可用的工作表列表"""
        if not self.file_input.value:
            return
            
        # 使用 BytesIO 包装字节数据
        first_file = BytesIO(self.file_input.value[0])
        xls = pd.ExcelFile(first_file)
        self.available_sheets = xls.sheet_names
        
        # 更新工作表选择器
        self.sheet_selector.options = self.available_sheets
        self.sheet_selector.value = []

    def _toggle_all_sheets(self, event):
        """切换全选/取消全选状态"""
        if set(self.sheet_selector.value) == set(self.available_sheets):
            self.sheet_selector.value = []
        else:
            self.sheet_selector.value = self.available_sheets

    def _merge_sheets(self, event):
        """合并选中的工作表"""
        if not self.file_input.value or not self.sheet_selector.value:
            print('请选择文件和工作表')
            return
            
        try:
            filepaths = self.file_input.value
            self.ttextract(filepaths, self.sheet_selector.value)
            print('合并完成！')
        except Exception as e:
            print(f'合并失败：{str(e)}')

    def ttextract(self, filepaths, sheets):
        self.all_merged_dfs = []
        
        with pd.ExcelWriter("表格汇总.xlsx") as writer:
            for sheet_name in sheets:
                self.all_dfs = []
                for i, file_bytes in enumerate(filepaths):
                    # 使用上传时的原始文件名
                    filename = self.file_input.filename[i]
                    print(f"读取文件 {filename} 的工作表 {sheet_name}")

                    # 使用 BytesIO 包装字节数据
                    excel_file = BytesIO(file_bytes)
                    try:
                        # 根据是否指定了单元格范围来读取数据
                        if self.start_range_input.value.strip():
                            start_cell = self.start_range_input.value.strip()
                            end_cell = self.end_range_input.value.strip() or start_cell
                            
                            # 提取起始和结束的行列
                            start_col = ''.join(filter(str.isalpha, start_cell)).upper()
                            end_col = ''.join(filter(str.isalpha, end_cell)).upper()
                            start_row = int(''.join(filter(str.isdigit, start_cell)))
                            end_row = int(''.join(filter(str.isdigit, end_cell)))
                            
                            # 计算列范围
                            usecols = None
                            if start_col and end_col:
                                from openpyxl.utils import column_index_from_string
                                start_col_idx = column_index_from_string(start_col) - 1
                                end_col_idx = column_index_from_string(end_col)
                                usecols = range(start_col_idx, end_col_idx)
                            
                            skiprows = start_row - 1
                            nrows = end_row - start_row + 1
                            
                            df = pd.read_excel(
                                excel_file,
                                sheet_name=sheet_name,
                                usecols=usecols,
                                skiprows=skiprows,
                                header=0 if self.use_header_checkbox.value else None,
                                dtype=str,
                                nrows=nrows
                            )
                        else:
                            df = pd.read_excel(
                                excel_file,
                                sheet_name=sheet_name,
                                header=0 if self.use_header_checkbox.value else None,
                                dtype=str
                            )
                        # 添加文件名和表名列
                        df.insert(0, "文件", filename)
                        df.insert(1, "表名", sheet_name)
                    except Exception as e:
                        print(f"处理文件 {filename} 时出错: {str(e)}")
                        continue

                    self.all_dfs.append(df)

                if self.all_dfs:
                    # 直接合并所有数据框，pandas会自动处理不同的列
                    self.merged_df = pd.concat(self.all_dfs, axis=0, ignore_index=True)
                    self.merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.all_merged_dfs.append(self.merged_df)
                else:
                    print(f"警告：没有成功读取任何数据到工作表 {sheet_name}")
            
            if self.all_merged_dfs:
                # 合并所有工作表的数据到总表
                final_merged_df = pd.concat(self.all_merged_dfs, axis=0, ignore_index=True)
                final_merged_df.to_excel(writer, sheet_name='合并总表', index=False)
                self.download_button.file = "表格汇总.xlsx"
                self.download_button.filename = "表格汇总.xlsx"
                self.download_button.visible = True

        print(f"已导出：表格汇总.xlsx")

# 使用示例
if __name__ == "__main__":
    merger = ExcelMerger()
    pn.serve(merger.layout)