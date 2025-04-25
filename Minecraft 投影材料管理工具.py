import sys
import os
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QTableView, 
                               QVBoxLayout, QWidget, QLabel, QHeaderView,
                               QDialog, QListWidget, QPushButton, QHBoxLayout,
                               QMessageBox, QFileDialog, QStyle, QStyleFactory,
                               QToolBar, QStyledItemDelegate, QItemDelegate,
                               QSizePolicy, QToolButton, QSpinBox, QMenu, QComboBox)
from PySide6.QtCore import Qt, QAbstractTableModel, QMimeData, QUrl, QSettings, QEvent, QSize, QSortFilterProxyModel, QTimer, QModelIndex
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QColor, QAction

# --- 添加 Minecraft 数量格式化函数 ---
def format_minecraft_quantity(value, unit, stack_size=64, box_slots=27):
    """将数值格式化为Minecraft单位字符串"""
    try:
        num_value = int(value) # 确保是整数
    except (ValueError, TypeError):
        return str(value) # 无法转换则返回原样

    if unit == "个" or num_value == 0:
        return str(num_value)

    elif unit == "组":
        stacks = num_value // stack_size
        items = num_value % stack_size
        if items == 0:
            return f"{stacks}组"
        elif stacks == 0:
             return f"{items}个" # 小于一组时只显示个数
        else:
            return f"{stacks}组 {items}个"

    elif unit == "盒":
        box_capacity = stack_size * box_slots
        boxes = num_value // box_capacity
        remainder_after_boxes = num_value % box_capacity
        stacks = remainder_after_boxes // stack_size
        items = remainder_after_boxes % stack_size

        parts = []
        if boxes > 0:
            parts.append(f"{boxes}盒")
        if stacks > 0:
            parts.append(f"{stacks}组")
        if items > 0:
            parts.append(f"{items}个")
        
        if not parts: # 如果刚好是0
             return "0"
             
        return " ".join(parts)

    else: # 未知单位，返回原值
        return str(num_value)
# --- 函数结束 ---

# 定义样式表
LIGHT_STYLE = """
QMainWindow, QDialog {
    background-color: #f5f5f5;
}
QTableView {
    border: 1px solid #d3d3d3;
    gridline-color: #f0f0f0;
    font-size: 10pt;
    background-color: white;
    alternate-background-color: #f9f9f9;
}
QHeaderView::section {
    background-color: #f0f0f0;
    padding: 5px;
    border: 1px solid #d3d3d3;
    font-weight: bold;
}
QTableView::item:selected {
    background-color: #cce8ff;
    color: black;
}
QTableView::item:focus {
    border: 2px solid #0078D7;
    background-color: #e5f1fb;
}
QPushButton {
    background-color: #0078D7;
    color: white;
    border: none;
    padding: 8px 16px;
    font-size: 10pt;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #1683d8;
}
QPushButton:pressed {
    background-color: #006cbe;
}
QLabel {
    color: #333333;
}
QToolBar {
    background-color: #f0f0f0;
    border-bottom: 1px solid #d3d3d3;
}
"""

DARK_STYLE = """
QMainWindow, QDialog {
    background-color: #2d2d2d;
}
QTableView {
    border: 1px solid #555555;
    gridline-color: #3a3a3a;
    font-size: 10pt;
    background-color: #2d2d2d;
    alternate-background-color: #353535;
    color: #e0e0e0;
}
QHeaderView::section {
    background-color: #3a3a3a;
    padding: 5px;
    border: 1px solid #555555;
    font-weight: bold;
    color: #e0e0e0;
}
QTableView::item:selected {
    background-color: #0c5f9e;
    color: white;
}
QTableView::item:focus {
    border: 2px solid #0078D7;
    background-color: #0c5f9e;
}
QPushButton {
    background-color: #0078D7;
    color: white;
    border: none;
    padding: 8px 16px;
    font-size: 10pt;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #1683d8;
}
QPushButton:pressed {
    background-color: #006cbe;
}
QLabel {
    color: #e0e0e0;
}
QStatusBar {
    color: #e0e0e0;
}
QToolBar {
    background-color: #3a3a3a;
    border-bottom: 1px solid #555555;
}
"""

class MaterialsModel(QAbstractTableModel):
    def __init__(self, data, file_path=None):
        super().__init__()
        self._data = data
        self._headers = ["材料名称", "总需求", "仍需收集", "已有数量", "收集状态", "快速完成"]
        # 保存文件路径
        self._file_path = file_path
        # 主题模式
        self._is_dark_mode = False
        # 添加排序支持
        self._sort_column = -1
        self._sort_order = Qt.AscendingOrder
        # --- 添加显示单位变量 --- 
        self._display_unit = "个" # 默认单位
        # --- 添加快速完成锁定状态 --- 
        self._quick_complete_locked = False
        
    def set_display_unit(self, unit):
        """设置当前显示单位"""
        if unit in ["个", "组", "盒"]:
            self._display_unit = unit
            # 注意：这里模型本身不触发 dataChanged，由调用者 (MainWindow) 负责
        
    def set_quick_complete_locked(self, locked):
        """设置快速完成列是否锁定"""
        if isinstance(locked, bool):
            self._quick_complete_locked = locked
            # 触发列5的flags和可能的显示变化
            num_rows = self.rowCount(QModelIndex()) # <-- 修正调用
            if num_rows > 0:
                col5_start_index = self.index(0, 5)
                col5_end_index = self.index(num_rows - 1, 5)
                # 发射所有相关角色以确保更新
                self.dataChanged.emit(col5_start_index, col5_end_index, [Qt.ItemFlags, Qt.CheckStateRole, Qt.DisplayRole, Qt.BackgroundRole])
                # 同时通知表头数据也可能需要更新（图标）
                self.headerDataChanged.emit(Qt.Horizontal, 5, 5)
        
    def data(self, index, role):
        if not index.isValid():
            return None
            
        if role == Qt.DisplayRole:
            row = index.row()
            col = index.column()
            
            # 安全检查
            if row < 0 or row >= len(self._data):
                return None
            
            value = self._data.iloc[row, col] # 获取原始值
            
            if col in [1, 2, 3]: # 总需求, 仍需收集, 已有数量
                 # 调用辅助函数进行格式化
                 # 确保 format_minecraft_quantity 函数存在且可调用
                 return format_minecraft_quantity(value, self._display_unit)
            elif col == 0: # 材料名称
                return str(value)
            elif col == 4: # 收集状态
                try:
                    missing = self._data.iloc[row, 2] # 依赖仍需收集列
                    missing_value = float(missing) if isinstance(missing, str) else missing
                    return "已完成" if missing_value == 0 else "未完成"
                except (ValueError, TypeError):
                    return "未完成"
            elif col == 5: # 快速完成列不显示文本
                 return ""
            else: # 其他列（如果有）
                 return str(value) # 其他列返回原始字符串
        
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        
        elif role == Qt.BackgroundRole:
            try:
                row = index.row()
                if row < 0 or row >= len(self._data):
                    return QColor("white") if not self._is_dark_mode else QColor("#2d2d2d")

                missing = self._data.iloc[row, 2]
                # 尝试将missing转换为数值进行比较
                missing_value = float(missing) if isinstance(missing, str) else missing
                if missing_value == 0:
                    # 根据主题返回不同的绿色
                    return QColor("#90EE90") if not self._is_dark_mode else QColor("#2E8B57")  # 浅绿色或深绿色
                else:
                    # 根据主题返回不同的背景色
                    return QColor("white") if not self._is_dark_mode else QColor("#2d2d2d")  # 白色或暗灰色
            except (ValueError, TypeError, IndexError):
                # 如果转换失败或索引错误，根据主题返回默认背景色
                return QColor("white") if not self._is_dark_mode else QColor("#2d2d2d")
        
        elif role == Qt.ForegroundRole and index.column() == 4:
            try:
                row = index.row()
                if row < 0 or row >= len(self._data):
                    return QColor("black") if not self._is_dark_mode else QColor("white")
                    
                # 为收集状态设置特殊文本颜色
                missing = self._data.iloc[row, 2]
                missing_value = float(missing) if isinstance(missing, str) else missing
                if missing_value == 0:
                    return QColor("darkgreen") if not self._is_dark_mode else QColor("#90EE90")
                else:
                    return QColor("darkred") if not self._is_dark_mode else QColor("#FF6B6B")
            except (ValueError, TypeError, IndexError) as e:
                return QColor("black") if not self._is_dark_mode else QColor("white")
        
        elif role == Qt.CheckStateRole and index.column() == 5:
            try:
                row = index.row()
                if row < 0 or row >= len(self._data):
                    return Qt.Unchecked
                    
                # 在"快速完成"列显示复选框
                missing = self._data.iloc[row, 2]
                missing_value = float(missing) if isinstance(missing, str) else missing
                return Qt.Checked if missing_value == 0 else Qt.Unchecked
            except (ValueError, TypeError, IndexError) as e:
                return Qt.Unchecked
                
        elif role == Qt.EditRole:
            row = index.row()
            col = index.column()
            if row < 0 or row >= len(self._data): return None

            if col == 3: # 只允许编辑"已有数量"
                return str(self._data.iloc[row, col]) # 返回原始值字符串
            # 其他列不允许直接编辑或返回其显示值
            elif col in [0, 1, 2, 4, 5]:
                 # 对于其他列，在编辑角色时，可以返回None或者显示值，但不应可编辑
                 # 返回显示值可能更一致，但要确保flags阻止编辑
                 return self.data(index, Qt.DisplayRole) # 返回格式化后的值，但flags应阻止编辑
            else:
                 return None
        
        return None
    
    def set_dark_mode(self, is_dark):
        """设置暗色模式状态"""
        self._is_dark_mode = is_dark
        # 通知视图刷新所有单元格
        self.layoutChanged.emit()
    
    def setData(self, index, value, role=Qt.EditRole):
        """处理数据编辑"""
        row = index.row()
        col = index.column()
        
        if role == Qt.EditRole and col == 3:  # 编辑"已有数量"列
            try:
                # 尝试将输入转换为整数
                available = int(value)
                if available < 0:
                    return False  # 不允许负数
                
                # 更新已有数量
                self._data.iloc[row, 3] = available
                
                # 更新仍需收集列
                total = int(self._data.iloc[row, 1])
                missing = max(0, total - available)
                self._data.iloc[row, 2] = missing
                
                # 发出数据改变信号，更新整行
                self.dataChanged.emit(
                    self.index(row, 0),  # 从第一列
                    self.index(row, 5)   # 到最后一列
                )
                
                # 保存更改到CSV文件 - 无论如何都尝试保存
                self.save_to_csv()
                
                return True
            except ValueError:
                return False
        
        elif role == Qt.CheckStateRole and col == 5:  # 切换"快速完成"复选框
            try:
                # 获取材料总需求量
                total = int(self._data.iloc[row, 1])
                
                if value == Qt.Checked:
                    # 如果勾选，将已有数量设置为总需求量
                    self._data.iloc[row, 3] = total
                    # 将仍需收集设置为0
                    self._data.iloc[row, 2] = 0
                else:
                    # 如果取消勾选，将已有数量设置为0
                    self._data.iloc[row, 3] = 0
                    # 将仍需收集设置为总需求量
                    self._data.iloc[row, 2] = total
                
                # 发出数据改变信号，更新整行
                self.dataChanged.emit(
                    self.index(row, 0),  # 从第一列
                    self.index(row, 5)   # 到最后一列
                )
                
                # 保存更改到CSV文件 - 无论如何都尝试保存
                self.save_to_csv()
                
                return True
            except Exception as e:
                return False
                
        return False
    
    def save_to_csv(self):
        """保存数据到CSV文件"""
        try:
            # 检查文件路径
            if not self._file_path:
                return False
                
            # 创建备份
            if os.path.exists(self._file_path):
                backup_path = f"{self._file_path}.bak"
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                os.rename(self._file_path, backup_path)
            
            # 保存数据
            self._data.to_csv(self._file_path, index=False)
            
            # 尝试显示保存成功提示
            try:
                # 获取主窗口对象
                for widget in QApplication.topLevelWidgets():
                    if isinstance(widget, MainWindow):
                        widget.show_notification("数据已保存", 2000)
                        break
            except Exception:
                pass
                
            return True
        except Exception:
            return False
    
    def rowCount(self, index):
        return len(self._data)
    
    def columnCount(self, index):
        return len(self._headers)
    
    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            # 返回存储在列表中的表头文本
            try:
                return self._headers[section]
            except IndexError:
                return None
        return None
        
    def flags(self, index):
        """返回项目标志"""
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        col = index.column()
        
        # "已有数量"列可编辑
        if col == 3:
            flags |= Qt.ItemIsEditable
        
        # "快速完成"列可勾选 (当未锁定时)
        if col == 5:
            is_locked = self._quick_complete_locked
            if not is_locked:
                 flags |= Qt.ItemIsUserCheckable
            
        return flags

    # 添加排序支持
    def sort(self, column, order):
        """按指定列和顺序排序数据"""
        self.layoutAboutToBeChanged.emit()
        
        # 记住当前排序列和顺序
        self._sort_column = column
        self._sort_order = order
        
        try:
            # 确定排序方向
            ascending = (order == Qt.AscendingOrder)
            
            # --- 如果是第5列，不执行排序 --- 
            if column == 5:
                self.layoutChanged.emit() # 仍然需要发射信号以防万一
                return
            # --- 结束处理第5列 ---
            
            if column == 0:  # 材料名称 - 字符串排序
                self._data = self._data.sort_values(by=self._data.columns[0], ascending=ascending)
            
            elif column == 1:  # 总需求 - 数值排序
                self._data["Total_num"] = pd.to_numeric(self._data["Total"], errors="coerce")
                self._data = self._data.sort_values(by="Total_num", ascending=ascending)
                self._data = self._data.drop(columns=["Total_num"])
            
            elif column == 2:  # 仍需收集 - 数值排序
                self._data["Missing_num"] = pd.to_numeric(self._data["Missing"], errors="coerce")
                self._data = self._data.sort_values(by="Missing_num", ascending=ascending)
                self._data = self._data.drop(columns=["Missing_num"])
            
            elif column == 3:  # 已有数量 - 数值排序
                self._data["Available_num"] = pd.to_numeric(self._data["Available"], errors="coerce")
                self._data = self._data.sort_values(by="Available_num", ascending=ascending)
                self._data = self._data.drop(columns=["Available_num"])
            
            elif column == 4:  # 收集状态 - 先按状态分组，再按Missing值排序
                self._data["IsComplete"] = self._data["Missing"].apply(lambda x: 0 if float(x) == 0 else 1)
                self._data["Missing_num"] = pd.to_numeric(self._data["Missing"], errors="coerce")
                self._data = self._data.sort_values(by=["IsComplete", "Missing_num"], ascending=ascending)
                self._data = self._data.drop(columns=["IsComplete", "Missing_num"])
                
            elif column == 5:  # 快速完成列 - 使用与收集状态相同的逻辑
                self._data["IsComplete"] = self._data["Missing"].apply(lambda x: 0 if float(x) == 0 else 1)
                self._data["Missing_num"] = pd.to_numeric(self._data["Missing"], errors="coerce")
                self._data = self._data.sort_values(by=["IsComplete", "Missing_num"], ascending=ascending)
                self._data = self._data.drop(columns=["IsComplete", "Missing_num"])
        
        except Exception as e:
            # 简单的备用排序方法
            try:
                col_name = self._data.columns[min(column, len(self._data.columns)-1)]
                self._data = self._data.sort_values(by=col_name, ascending=ascending)
            except:
                pass
        
        self.layoutChanged.emit()

class FileSelectDialog(QDialog):
    def __init__(self, file_list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择CSV文件")
        self.setFixedSize(500, 400)
        self.selected_file = None
        
        # 设置Windows风格
        self.setStyle(QStyleFactory.create("WindowsVista"))
        
        # 创建布局
        layout = QVBoxLayout()
        
        # 添加指导标签
        label = QLabel("请选择要加载的CSV文件:")
        label.setStyleSheet("font-size: 12pt; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(label)
        
        # 创建文件列表
        self.file_list_widget = QListWidget()
        for file in file_list:
            self.file_list_widget.addItem(file)
        self.file_list_widget.setStyleSheet("font-size: 10pt;")
        layout.addWidget(self.file_list_widget)
        
        # 创建按钮区域
        button_layout = QHBoxLayout()
        
        self.select_button = QPushButton("选择")
        self.select_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D7;
                color: white;
                border: none;
                padding: 8px 16px;
                font-size: 10pt;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1683d8;
            }
            QPushButton:pressed {
                background-color: #006cbe;
            }
        """)
        self.select_button.clicked.connect(self.accept_selection)
        
        self.cancel_button = QPushButton("取消")
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #e1e1e1;
                border: none;
                padding: 8px 16px;
                font-size: 10pt;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d1d1d1;
            }
            QPushButton:pressed {
                background-color: #c1c1c1;
            }
        """)
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def accept_selection(self):
        selected_items = self.file_list_widget.selectedItems()
        if selected_items:
            self.selected_file = selected_items[0].text()
            self.accept()
        else:
            QMessageBox.warning(self, "警告", "请先选择一个文件!")

class CheckBoxDelegate(QStyledItemDelegate):
    """自定义复选框委托，处理点击事件"""
    def __init__(self, parent=None):
        super().__init__(parent)
    
    def paint(self, painter, option, index):
        # 使用默认绘制方法
        super().paint(painter, option, index)
    
    def editorEvent(self, event, model, option, index):
        # 只处理"快速完成"列
        if index.column() != 5:
            return False
            
        # --- 检查锁定状态 --- 
        # 需要访问源模型来获取锁定状态
        actual_model = model # 默认是代理模型
        is_locked = False # 默认未锁定
        if isinstance(model, QSortFilterProxyModel):
             source_model = model.sourceModel()
             if source_model and hasattr(source_model, '_quick_complete_locked'):
                 is_locked = source_model._quick_complete_locked
             else:
                 is_locked = True # 无法获取状态时，安全起见视为锁定
        elif hasattr(model, '_quick_complete_locked'): # 检查基本模型
             is_locked = model._quick_complete_locked
        else:
             is_locked = True # 无法获取状态时，安全起见视为锁定
             
        if is_locked:
            return False # 如果锁定了，阻止事件处理
        # --- 结束检查 --- 
            
        # 处理鼠标点击事件 (只有未锁定时才会执行到这里)
        if event.type() == QEvent.MouseButtonRelease:
            # 获取当前复选框状态
            # 注意：这里用 model (可能是代理) 来获取当前显示的状态
            checkState = index.data(Qt.CheckStateRole)
            
            # 切换状态
            newState = Qt.Unchecked if checkState == Qt.Checked else Qt.Checked
            
            # 设置新状态 (通过 model，它会路由到源模型)
            return model.setData(index, newState, Qt.CheckStateRole)
            
        return False

# 添加过滤代理模型
class MaterialFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.hide_collected = False
    
    # Override the sort method
    def sort(self, column, order):
        """Overrides the default sort to ensure the source model's sort is called 
           and the proxy signals layout change correctly.
        """
        source_model = self.sourceModel()
        if source_model:
            # Store current sort column and order in the proxy using correct method names
            self.sortColumn = column # Set the sort column directly
            self.sortOrder = order   # Set the sort order directly
            
            # Tell the source model to sort itself (this will trigger the prints)
            source_model.sort(column, order)
            
            # Emit layoutChanged from the proxy model to update the view
            self.layoutChanged.emit()
        else:
            # Fallback to default behavior if no source model
            super().sort(column, order)

    def filterAcceptsRow(self, source_row, source_parent):
        """根据过滤设置决定是否显示行"""
        if not self.hide_collected:
            return True  # 如果不隐藏已收集的，显示所有行
        
        try:
            # 安全检查
            source_model = self.sourceModel()
            if source_model is None:
                return True
                
            if source_row < 0 or source_row >= source_model.rowCount(source_parent):
                return True
            
            # 检查数据的Missing列（索引2），如果等于0表示已收集完成
            missing_index = source_model.index(source_row, 2, source_parent)
            if not missing_index.isValid():
                return True
                
            missing_value = source_model.data(missing_index, Qt.DisplayRole)
            if missing_value is None:
                return True
                
            # 由于missing_value是字符串类型，需要转换为数值比较
            try:
                # 尝试将missing_value转换为数值进行比较
                missing_num = float(missing_value) if isinstance(missing_value, str) else missing_value
                return missing_num != 0
            except (ValueError, TypeError):
                # 如果转换失败，默认显示该行
                return True
        except Exception as e:
            return True  # 出错时默认显示该行
    
    def set_hide_collected(self, hide):
        """设置是否隐藏已收集项"""
        try:
            if self.hide_collected != hide:
                self.hide_collected = hide
                self.invalidateFilter()  # 刷新过滤器
        except Exception as e:
            self.hide_collected = False  # 出错时重置为显示所有项

class MainWindow(QMainWindow):
    def __init__(self, csv_data=None, csv_path=None):
        super().__init__()
        
        self.setWindowTitle("Minecraft 投影材料管理")
        self.setGeometry(100, 100, 800, 600)
        
        # 改进设置加载，指定完整的组织名和应用名
        self.settings = QSettings("MinecraftTool", "MaterialManager")
        
        # 加载设置并打印以便调试
        self.is_dark_mode = self.settings.value("is_dark_mode", False, type=bool)
        self.is_always_on_top = self.settings.value("is_always_on_top", False, type=bool)
        self.is_hiding_collected = self.settings.value("is_hiding_collected", False, type=bool)

        # 每次启动时总是将投影数量设为1，忽略保存的值
        self.projection_count = 1
        # 保存设置以保持一致性
        self.settings.setValue("projection_count", 1)
        self.settings.sync()
        
        # 保存当前数据和路径
        self.original_csv_data = csv_data  # 原始数据（单个投影）
        self.current_csv_data = self.apply_projection_multiplier(csv_data, self.projection_count) if csv_data is not None else None
        self.current_csv_path = csv_path
        
        # 设置Windows风格
        self.setStyle(QStyleFactory.create("WindowsVista"))
        
        # 应用窗口置顶状态
        self.set_always_on_top(self.is_always_on_top)
        
        # 启用接收拖放
        self.setAcceptDrops(True)
        
        # 添加右下角提示标签
        self.save_notification = QLabel("", self)
        self.save_notification.setStyleSheet("""
            background-color: rgba(0, 120, 215, 0.8);
            color: white;
            border-radius: 4px;
            padding: 8px;
            font-weight: bold;
        """)
        self.save_notification.setAlignment(Qt.AlignCenter)
        self.save_notification.hide()
        
        # 创建显示提示的计时器
        self.notification_timer = QTimer(self)
        self.notification_timer.timeout.connect(self.hide_notification)
        
        # 添加用于跟踪当前排序状态的变量
        self._current_sort_column = -1
        self._current_sort_order = Qt.AscendingOrder # 默认初始升序
        
        # --- 添加当前显示单位变量 --- 
        self.current_unit = self.settings.value("display_unit", "个", type=str) # 默认是"个"
        # --- 添加快速完成锁定状态 --- 
        self.quick_complete_locked = False # 初始为未锁定
        # --- 结束添加 --- 
        
        # 创建工具栏
        self.toolbar = QToolBar()
        self.toolbar.setMovable(False)
        self.addToolBar(self.toolbar)
        
        # 添加主题切换按钮
        self.theme_action = QAction("切换暗色模式" if not self.is_dark_mode else "切换亮色模式", self)
        self.theme_action.triggered.connect(self.toggle_theme)
        self.toolbar.addAction(self.theme_action)
        
        # 添加忽略已收集项按钮
        self.hide_collected_action = QAction(
            "显示已收集材料" if self.is_hiding_collected else "隐藏已收集材料", 
            self
        )
        self.hide_collected_action.triggered.connect(self.toggle_hide_collected)
        self.toolbar.addAction(self.hide_collected_action)
        
        # 添加重置列宽按钮
        self.reset_width_action = QAction("重置列宽", self)
        self.reset_width_action.triggered.connect(self.reset_column_widths)
        self.toolbar.addAction(self.reset_width_action)
        
        # 添加投影数量选择器
        projection_label = QLabel("投影数量:")
        self.toolbar.addWidget(projection_label)
        
        self.projection_spinner = QSpinBox()
        self.projection_spinner.setMinimum(1)
        self.projection_spinner.setMaximum(10)  # 限制最多10个投影
        self.projection_spinner.setValue(self.projection_count)
        self.projection_spinner.valueChanged.connect(self.change_projection_count)
        self.toolbar.addWidget(self.projection_spinner)
        
        # --- 添加单位选择器 --- 
        unit_label = QLabel("  显示单位:") # 加点空格美观
        self.toolbar.addWidget(unit_label)
        
        self.unit_combobox = QComboBox()
        self.unit_combobox.addItems(["个", "组", "盒"])
        # 设置初始值
        current_index = self.unit_combobox.findText(self.current_unit)
        if current_index != -1:
            self.unit_combobox.setCurrentIndex(current_index)
        # 连接信号
        self.unit_combobox.currentIndexChanged.connect(self.change_display_unit)
        self.toolbar.addWidget(self.unit_combobox)
        # --- 结束添加 --- 
        
        # 添加弹簧部件将后续控件推到右边
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.toolbar.addWidget(spacer)
        
        # 添加窗口置顶按钮（在右侧）
        self.topmost_action = QAction(
            "取消置顶窗口" if self.is_always_on_top else "置顶窗口", 
            self
        )
        self.topmost_action.triggered.connect(self.toggle_always_on_top)
        self.toolbar.addAction(self.topmost_action)
        
        # 创建主部件和布局
        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.setCentralWidget(self.main_widget)
        
        # 应用当前主题
        self.apply_theme()
        
        # 初始化界面
        self.initialize_interface()
        
    def apply_projection_multiplier(self, data, multiplier):
        """根据投影数量调整材料需求量"""
        if data is None or data.empty:
            return data
            
        # 创建副本，避免修改原始数据
        result = data.copy()
        
        # 调整总需求量和仍需收集量
        if 'Total' in result.columns:
            result['Total'] = result['Total'] * multiplier
        
        # 根据已有数量重新计算仍需收集量
        if 'Total' in result.columns and 'Available' in result.columns:
            result['Missing'] = result['Total'] - result['Available']
            # 确保Missing不小于0
            result['Missing'] = result['Missing'].apply(lambda x: max(0, x))
            
        return result
        
    def change_projection_count(self, count):
        """更改投影数量，保留已有数量和当前排序状态"""
        if count == self.projection_count:
            return

        # --- Essential Check: Ensure model and data exist ---
        if not hasattr(self, 'model') or self.model is None or self.original_csv_data is None:
            self.statusBar().showMessage("错误：无法更改投影数量，模型或数据未加载", 3000)
            # Reset spinner value if possible
            if hasattr(self, 'projection_spinner'):
                self.projection_spinner.blockSignals(True)
                self.projection_spinner.setValue(self.projection_count)
                self.projection_spinner.blockSignals(False)
            return
        # --- End Check ---

        try:
            # --- Modify data directly in the source model ---
            source_model = self.model # Access the existing source model
            # Create a dictionary of original totals for faster lookup
            original_totals = self.original_csv_data.set_index('Item')['Total'].to_dict()

            # Iterate through the rows of the model's data
            # Important: We are modifying the DataFrame held by the model (_data)
            num_rows = source_model.rowCount(None) # Get row count from model
            for i in range(num_rows):
                # Get item name using model's index method for safety
                item_name_index = source_model.index(i, 0) # Column 0 is Item
                item_name = source_model.data(item_name_index, Qt.DisplayRole)

                if item_name is None: continue # Skip if item name couldn't be retrieved

                # Get original total
                original_total = original_totals.get(item_name, 0)
                # Calculate new total
                new_total = original_total * count

                # Get current available quantity directly from model data
                available_index = source_model.index(i, 3) # Column 3 is Available
                current_available_str = source_model.data(available_index, Qt.EditRole) # Use EditRole or DisplayRole
                try:
                    current_available = int(current_available_str)
                except (ValueError, TypeError, TypeError):
                     current_available = 0

                # Calculate new missing quantity
                new_missing = max(0, new_total - current_available)

                # --- Update the model's internal DataFrame ---
                # Use .iloc for direct access if _data is guaranteed to be the DataFrame
                try:
                    source_model._data.iloc[i, 1] = new_total   # Update Total (Column 1)
                    source_model._data.iloc[i, 2] = new_missing # Update Missing (Column 2)
                    # Available (Column 3) remains unchanged
                except IndexError:
                     continue # Should not happen if rowCount is correct
                # --- End DataFrame Update ---

            # --- Finished modifying data ---

            # --- Signal Data Change (Crucial for View Update) ---
            # Signal changes for the affected columns (Total and Missing) for all rows
            start_index_col1 = source_model.index(0, 1)
            end_index_col2 = source_model.index(num_rows - 1, 2)
            source_model.dataChanged.emit(start_index_col1, end_index_col2, [Qt.DisplayRole, Qt.EditRole])
            # Also signal change for '收集状态' as it depends on 'Missing'
            start_index_col4 = source_model.index(0, 4)
            end_index_col4 = source_model.index(num_rows - 1, 4)
            source_model.dataChanged.emit(start_index_col4, end_index_col4, [Qt.DisplayRole, Qt.BackgroundRole, Qt.ForegroundRole])
            # Signal change for '快速完成' checkbox state
            start_index_col5 = source_model.index(0, 5)
            end_index_col5 = source_model.index(num_rows - 1, 5)
            source_model.dataChanged.emit(start_index_col5, end_index_col5, [Qt.CheckStateRole])
            # --- End Signal ---

            # Update the projection count variable and settings
            self.projection_count = count
            self.settings.setValue("projection_count", count)
            self.settings.sync()

            # --- Re-apply Sort Order ---
            if self._current_sort_column != -1:
                 # Re-sort the proxy model to maintain the current sort order visually
                 if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                      self.proxy_model.sort(self._current_sort_column, self._current_sort_order)
                 elif hasattr(self, 'model'): # Fallback if no proxy
                      self.model.sort(self._current_sort_column, self._current_sort_order)
                 # Ensure header indicator is updated
                 if hasattr(self, 'table_view'):
                      self.table_view.horizontalHeader().setSortIndicator(self._current_sort_column, self._current_sort_order)
            # --- End Re-apply Sort ---

            # Update statistics display
            self.update_statistics()

            # Show notification
            self.show_notification(f"已更新为 {count} 个投影的材料需求", 3000)

            # Save the modified data
            source_model.save_to_csv()

        except Exception as e:
            QMessageBox.warning(self, "错误", f"更改投影数量时出错: {e}")
            # Attempt to reset spinner value
            if hasattr(self, 'projection_spinner'):
                self.projection_spinner.blockSignals(True)
                self.projection_spinner.setValue(self.projection_count)
                self.projection_spinner.blockSignals(False)

    def initialize_interface(self):
        """初始化用户界面"""
        try:
            # 清除现有布局中的所有部件
            self.clear_layout()

            # 确保投影数量选择器的值与当前设置一致
            if hasattr(self, 'projection_spinner'):
                self.projection_spinner.setValue(self.projection_count)
        
            # 添加标题
            title = QLabel("Minecraft 投影材料清单")
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("font-size: 20px; font-weight: bold;")
            self.main_layout.addWidget(title)

            # 显示当前加载的文件
            if self.current_csv_path:
                file_label = QLabel(f"当前文件: {os.path.basename(self.current_csv_path)}")
                file_label.setAlignment(Qt.AlignCenter)
                file_label.setStyleSheet("font-size: 10px; color: gray;")
                self.main_layout.addWidget(file_label)
            else:
                # 添加拖放提示和使用指南
                self.add_empty_state_guidance()

            if self.current_csv_data is not None and not self.current_csv_data.empty:
                self.initialize_table_view()
            else:
                # 如果没有数据，显示提示信息
                no_data_label = QLabel("未能加载有效的CSV数据，请检查文件格式")
                no_data_label.setAlignment(Qt.AlignCenter)
                error_color = "red" if not self.is_dark_mode else "#ff6b6b"
                no_data_label.setStyleSheet(f"font-size: 14px; color: {error_color}; margin: 20px;")
                self.main_layout.addWidget(no_data_label)

            # 添加加载按钮和拖放提示
            self.add_load_button_and_hint()

            # 更新窗口标题
            if self.current_csv_path:
                self.setWindowTitle(f"Minecraft 投影材料管理 - {os.path.basename(self.current_csv_path)}")
                # 文件加载成功，状态栏显示消息
                status_message = f"已成功加载文件: {os.path.basename(self.current_csv_path)}"
                self.statusBar().showMessage(status_message, 3000)
        except Exception:
            # 确保至少有一个基本界面
            self.show_error_interface("初始化界面出错")
            
    def clear_layout(self):
        """清除布局中的所有部件"""
        try:
            while self.main_layout.count():
                item = self.main_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
        except Exception:
            pass
            
    def add_empty_state_guidance(self):
        """添加空状态下的指导信息"""
        info_container = QWidget()
        info_layout = QVBoxLayout(info_container)
        
        # 主提示
        drop_hint = QLabel("未检测到CSV材料文件")
        drop_hint.setAlignment(Qt.AlignCenter)
        drop_hint.setStyleSheet("font-size: 16px; color: #444; font-weight: bold; margin: 10px;")
        info_layout.addWidget(drop_hint)
        
        # 方法一：拖放
        method1 = QLabel("方法一：将CSV文件拖放到此窗口")
        method1.setAlignment(Qt.AlignCenter)
        method1.setStyleSheet("font-size: 14px; color: #0078D7; margin-top: 20px;")
        info_layout.addWidget(method1)
        
        # 方法二：选择文件
        method2 = QLabel("方法二：点击下方按钮选择CSV文件")
        method2.setAlignment(Qt.AlignCenter)
        method2.setStyleSheet("font-size: 14px; color: #0078D7; margin-top: 10px;")
        info_layout.addWidget(method2)
        
        # 方法三：放入同目录
        method3 = QLabel("方法三：将CSV文件放入程序所在目录，然后重启程序")
        method3.setAlignment(Qt.AlignCenter)
        method3.setStyleSheet("font-size: 14px; color: #0078D7; margin-top: 10px;")
        info_layout.addWidget(method3)
        
        # 示例格式
        format_hint = QLabel("CSV文件格式示例：材料名称,总需求量,仍需收集量,已有数量")
        format_hint.setAlignment(Qt.AlignCenter)
        format_hint.setStyleSheet("font-size: 12px; color: #666; font-style: italic; margin-top: 30px;")
        info_layout.addWidget(format_hint)
        
        self.main_layout.addWidget(info_container)
            
    def initialize_table_view(self):
        """初始化表格视图"""
        try:
            # 创建表格视图
            self.table_view = QTableView()

            # 创建数据模型
            self.model = MaterialsModel(self.current_csv_data, self.current_csv_path)

            self.model.set_dark_mode(self.is_dark_mode)
            # 通知模型当前的显示单位
            self.model.set_display_unit(self.current_unit)

            # 安全地设置过滤代理模型
            try:
                self.proxy_model = MaterialFilterProxyModel(self)
                self.proxy_model.setSourceModel(self.model)
                self.proxy_model.set_hide_collected(self.is_hiding_collected)
                self.table_view.setModel(self.proxy_model)
            except Exception:
                self.table_view.setModel(self.model)
                self.proxy_model = None

            # 设置自定义委托处理复选框交互
            try:
                self.checkbox_delegate = CheckBoxDelegate(self.table_view)
                self.table_view.setItemDelegateForColumn(5, self.checkbox_delegate)
            except Exception:
                pass

            # 启用编辑模式
            self.table_view.setEditTriggers(QTableView.DoubleClicked | QTableView.SelectedClicked)
        
        # 设置表格属性
            self.configure_table_view()

            # 连接数据更改信号
            self.model.dataChanged.connect(self.update_statistics)
            if self.proxy_model:
                self.model.dataChanged.connect(self.refresh_filter)

            self.main_layout.addWidget(self.table_view)

            # 添加编辑提示
            self.add_edit_hint()

            # 添加统计信息
            self.stats_label = QLabel()
            self.update_statistics()
            self.stats_label.setAlignment(Qt.AlignCenter)
            self.stats_label.setStyleSheet("font-size: 12px; margin-top: 10px;")
            self.main_layout.addWidget(self.stats_label)
        except Exception:
            error_label = QLabel("加载表格出错")
            error_label.setAlignment(Qt.AlignCenter)
            error_label.setStyleSheet("font-size: 14px; color: red; margin: 20px;")
            self.main_layout.addWidget(error_label)
            
    def configure_table_view(self):
        """配置表格视图属性"""
        try:
            header = self.table_view.horizontalHeader()

            # 先设置所有列为Fixed，再单独设置特定列的模式
            for i in range(6):  # 总共6列
                header.setSectionResizeMode(i, QHeaderView.Fixed)

            # 设置各列的初始宽度 - 使总宽度接近800像素（窗口宽度）
            window_width = self.width()  # 获取当前窗口宽度
            total_width = window_width - 34  # 右侧边距设为34像素

            # 根据比例分配宽度
            self.table_view.setColumnWidth(0, int(total_width * 0.38))  # 材料名称列约38%宽度
            self.table_view.setColumnWidth(1, int(total_width * 0.12))  # 总需求约12%宽度
            self.table_view.setColumnWidth(2, int(total_width * 0.12))  # 仍需收集约12%宽度
            self.table_view.setColumnWidth(3, int(total_width * 0.12))  # 已有数量约12%宽度
            self.table_view.setColumnWidth(4, int(total_width * 0.13))  # 收集状态约13%宽度
            self.table_view.setColumnWidth(5, int(total_width * 0.13))  # 快速完成约13%宽度

            # 启用用户调整列宽
            header.setSectionResizeMode(0, QHeaderView.Interactive)  # 材料名称可调整
            header.setSectionResizeMode(1, QHeaderView.Interactive)  # 总需求可调整
            header.setSectionResizeMode(2, QHeaderView.Interactive)  # 仍需收集可调整
            header.setSectionResizeMode(3, QHeaderView.Interactive)  # 已有数量可调整
            header.setSectionResizeMode(4, QHeaderView.Interactive)  # 收集状态可调整
            header.setSectionResizeMode(5, QHeaderView.Interactive)  # 快速完成可调整

            # 设置拖动调整行为 - 禁用最后一列自动拉伸
            header.setStretchLastSection(False)  # 最后一列不自动拉伸

            # 启用排序功能 (主要用于显示指示器)
            self.table_view.setSortingEnabled(True)
            header.setSortIndicatorShown(True)  # 显示排序箭头
            header.setHighlightSections(True)   # 高亮排序列

            # 启用表头点击
            header.setSectionsClickable(True)
            # 连接单击信号到新的处理函数
            header.sectionClicked.connect(self.handle_header_click)

            # --- 设置初始的快速完成列 Tooltip --- 
            initial_tooltip = "点击锁定/解锁此列功能 (当前已解锁)" # 默认是解锁
            # QHeaderView 没有直接设置单个 section tooltip 的标准方法
            # 我们将 tooltip 设置到整个 header 上，并在点击时更新它
            header.setToolTip(initial_tooltip)
            # --- 结束设置 --- 

            self.table_view.verticalHeader().setVisible(False)
            self.table_view.setAlternatingRowColors(True)
        except Exception:
            pass
            
    def add_edit_hint(self):
        """添加编辑提示"""
        # 更新提示文本，加入右键排序的提示
        edit_hint_text = "提示: 1.双击\"已有数量\"列可以直接修改数值；2.单击表头可以排序。"
        edit_hint = QLabel(edit_hint_text)
        edit_hint.setAlignment(Qt.AlignCenter)
        # 根据主题调整提示颜色
        hint_color = "#0078D7" if not self.is_dark_mode else "#4da6ff"
        edit_hint.setStyleSheet(f"font-size: 11px; color: {hint_color}; margin-top: 5px;")
        self.main_layout.addWidget(edit_hint)
        
    def add_load_button_and_hint(self):
        """添加加载按钮和拖放提示"""
        self.load_button = QPushButton("加载其他CSV文件")
        self.load_button.clicked.connect(self.load_csv_file)
        self.main_layout.addWidget(self.load_button, alignment=Qt.AlignCenter)
        
        # 添加拖放提示
        drop_hint = QLabel("或将CSV文件拖放至此处")
        drop_hint.setAlignment(Qt.AlignCenter)
        hint_color = "#666" if not self.is_dark_mode else "#aaa"
        drop_hint.setStyleSheet(f"font-size: 10px; color: {hint_color}; font-style: italic;")
        self.main_layout.addWidget(drop_hint)
        
    def show_error_interface(self, error_message):
        """显示错误界面"""
        self.clear_layout()
        
        error_label = QLabel(f"发生错误: {error_message}")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setStyleSheet("font-size: 14px; color: red; margin: 20px;")
        self.main_layout.addWidget(error_label)
        
        retry_button = QPushButton("重新加载界面")
        retry_button.clicked.connect(self.initialize_interface)
        self.main_layout.addWidget(retry_button, alignment=Qt.AlignCenter)
        
        load_button = QPushButton("加载其他CSV文件")
        load_button.clicked.connect(self.load_csv_file)
        self.main_layout.addWidget(load_button, alignment=Qt.AlignCenter)
            
    def toggle_theme(self):
        """切换亮色/暗色主题"""
        try:
            # 记录当前状态
            old_is_dark_mode = self.is_dark_mode
            old_is_hiding_collected = self.is_hiding_collected
            column_widths = []
            
            # 在切换主题前保存当前列宽（如果表格存在）
            if hasattr(self, 'table_view'):
                for col in range(6):  # 假设有6列
                    column_widths.append(self.table_view.columnWidth(col))
            
            # 临时禁用过滤，避免主题切换与过滤器冲突
            if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                self.proxy_model.set_hide_collected(False)
            
            # 切换主题状态
            self.is_dark_mode = not self.is_dark_mode
            
            # 保存设置
            self.settings.setValue("is_dark_mode", self.is_dark_mode)
            
            # 更新主题动作文本
            self.theme_action.setText("切换亮色模式" if self.is_dark_mode else "切换暗色模式")
            
            # 应用新主题样式
            style = DARK_STYLE if self.is_dark_mode else LIGHT_STYLE
            self.setStyleSheet(style)
            
            # 重新创建界面 - 但是保留当前数据和设置
            current_data = self.current_csv_data
            current_path = self.current_csv_path
            
            # 清空布局
            self.clear_layout()
            
            # 重新初始化界面
            self.initialize_interface()
            
            # 如果表格存在，恢复列宽
            if hasattr(self, 'table_view') and column_widths:
                for col, width in enumerate(column_widths):
                    if col < 6:  # 确保不会越界
                        self.table_view.setColumnWidth(col, width)
            
            # 恢复过滤状态
            if old_is_hiding_collected and hasattr(self, 'proxy_model') and self.proxy_model is not None:
                # 确保隐藏状态一致
                self.is_hiding_collected = old_is_hiding_collected
                self.hide_collected_action.setText(
                    "显示已收集材料" if self.is_hiding_collected else "隐藏已收集材料"
                )
                # 延迟应用过滤器，避免在界面刷新过程中应用
                QApplication.processEvents()
                self.proxy_model.set_hide_collected(self.is_hiding_collected)
                self.update_statistics()
                
            # 显示切换主题成功提示
            theme_msg = "已切换为暗色模式" if self.is_dark_mode else "已切换为亮色模式"
            self.show_notification(theme_msg, 2000)
        except Exception:
            QMessageBox.warning(self, "错误", "切换主题时出错")
            # 重置为安全状态
            self.is_dark_mode = old_is_dark_mode
            style = DARK_STYLE if self.is_dark_mode else LIGHT_STYLE
            self.setStyleSheet(style)
            self.theme_action.setText("切换亮色模式" if self.is_dark_mode else "切换暗色模式")
    
    def apply_theme(self):
        """应用当前主题"""
        try:
            # 应用全局样式表
            style = DARK_STYLE if self.is_dark_mode else LIGHT_STYLE
            self.setStyleSheet(style)
        except Exception:
            pass
    
    def toggle_hide_collected(self):
        """切换是否隐藏已收集项"""
        try:
            self.is_hiding_collected = not self.is_hiding_collected

            # 保存设置
            self.settings.setValue("is_hiding_collected", self.is_hiding_collected)

            # 更新按钮文本
            self.hide_collected_action.setText(
                "显示已收集材料" if self.is_hiding_collected else "隐藏已收集材料"
            )

            # 使用安全的方式应用过滤
            self.apply_filter()

            # 显示切换状态成功提示
            hide_msg = "已隐藏已收集材料" if self.is_hiding_collected else "已显示已收集材料"
            self.show_notification(hide_msg, 2000)
        except Exception:
            QMessageBox.warning(self, "错误", "切换显示状态时出错")
            # 重置状态并重试
            self.is_hiding_collected = False
            self.hide_collected_action.setText("隐藏已收集材料")
            try:
                self.initialize_interface()
            except:
                pass
    
    def apply_filter(self):
        """安全地应用过滤器"""
        try:
            if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                self.proxy_model.set_hide_collected(self.is_hiding_collected)
                self.update_statistics()
            else:
                # 如果代理模型不存在，重新初始化界面
                self.initialize_interface()
        except Exception:
            # 出错时重置界面
            self.initialize_interface()
    
    def refresh_filter(self):
        """刷新过滤器，用于数据变化后更新显示"""
        try:
            if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                self.proxy_model.invalidateFilter()
        except Exception:
            pass
    
    def load_and_update_csv(self, file_path):
        """加载CSV文件并更新界面"""
        try:
            # 加载新的CSV文件
            original_data = load_csv(file_path)
            
            # 保存原始数据
            self.original_csv_data = original_data
            
            # 重置投影数量为1
            self.projection_count = 1
            self.settings.setValue("projection_count", 1)
            self.settings.sync()
            if hasattr(self, 'projection_spinner'):
                self.projection_spinner.setValue(1)
            
            # 应用投影数量乘数 (现在是1)
            self.current_csv_data = self.apply_projection_multiplier(original_data, self.projection_count)
            
            # 更新文件路径
            self.current_csv_path = file_path
            
            # 初始化界面
            self.initialize_interface()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载CSV文件时出错: {str(e)}")
            self.show_error_interface(str(e))
    
    def update_statistics(self):
        """更新统计信息"""
        try:
            if hasattr(self, 'model') and hasattr(self, 'stats_label'):
                # 获取原始数据的总数
                total_items = len(self.model._data)
                
                # 获取已完成项目数量
                missing_col = 2  # Missing是第三列(索引2)
                completed_items = 0 # 初始化为0
                try:
                    # 使用安全的方式获取
                    completed_count = 0
                    for idx in range(len(self.model._data)):
                        try:
                            missing_value = self.model._data.iloc[idx, missing_col]
                            # 尝试转换为数值
                            missing_num = float(missing_value) if isinstance(missing_value, str) else missing_value
                            if missing_num == 0:
                                completed_count += 1
                        except:
                            pass
                    completed_items = completed_count
                except:
                    # 如果在获取过程中出错，completed_items 保持 0
                    pass 
        
                # 计算进度百分比
                completion_percent = (completed_items / total_items) * 100 if total_items > 0 else 0
        
                # 获取当前显示的行数
                visible_items = 0
                if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                    try:
                        visible_items = self.proxy_model.rowCount()
                    except:
                        visible_items = total_items # 出错时假设所有行都可见
                else:
                    visible_items = total_items
                    
                hidden_items = total_items - visible_items
                
                # 更新标签文本
                stats_text = f"材料总数: {total_items} | 已完成: {completed_items} | 进度: {completion_percent:.1f}%"
                
                # 如果有隐藏项目，显示信息
                if hidden_items > 0:
                    stats_text += f" | 已隐藏: {hidden_items}项"
                    
                self.stats_label.setText(stats_text)
        except Exception:
            # 出错时显示简单统计
            if hasattr(self, 'stats_label'):
                self.stats_label.setText("统计信息加载出错，请重新加载文件")
    
    def load_csv_file(self):
        # 使用文件对话框选择CSV文件
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择CSV文件", "", "CSV文件 (*.csv)"
        )
        
        if file_path:
            self.load_and_update_csv(file_path)
    
    def toggle_always_on_top(self):
        """切换窗口置顶状态"""
        self.is_always_on_top = not self.is_always_on_top
        
        # 保存设置
        self.settings.setValue("is_always_on_top", self.is_always_on_top)
        
        # 更新按钮文本
        self.topmost_action.setText("取消置顶窗口" if self.is_always_on_top else "置顶窗口")
        
        # 设置窗口置顶状态
        self.set_always_on_top(self.is_always_on_top)
    
    def set_always_on_top(self, enable):
        """设置窗口是否置顶"""
        # 保存当前窗口状态
        is_visible = self.isVisible()
        
        # 获取当前窗口标志，但重新设置基本窗口类型
        # 确保窗口始终具有基本窗口类型标志和系统按钮
        flags = Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowSystemMenuHint
        
        # 添加或移除置顶标志
        if enable:
            flags |= Qt.WindowStaysOnTopHint
        
        # 设置新的窗口标志
        self.setWindowFlags(flags)
        
        # 如果窗口之前是可见的，重新显示它
        if is_visible:
            self.show()

    def dragEnterEvent(self, event: QDragEnterEvent):
        """处理拖入文件的事件"""
        try:
            # 检查是否包含URL（文件路径）
            if event.mimeData().hasUrls():
                # 获取第一个文件的URL
                url = event.mimeData().urls()[0]
                # 检查是否为本地文件且为CSV文件
                if url.isLocalFile() and url.toLocalFile().lower().endswith('.csv'):
                    event.acceptProposedAction()
                    return
            # 如果不符合条件，不接受拖放
            event.ignore()
        except Exception:
            event.ignore()

    def dragMoveEvent(self, event):
        """处理拖动移动事件"""
        try:
            # 如果已经在dragEnterEvent中接受，这里也接受
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
        except Exception:
            event.ignore()
    
    def dropEvent(self, event: QDropEvent):
        """处理文件放下的事件"""
        try:
            # 获取第一个文件的路径
            file_path = event.mimeData().urls()[0].toLocalFile()
            
            # 检查是否为CSV文件
            if file_path.lower().endswith('.csv'):
                # 加载CSV文件
                self.load_and_update_csv(file_path)
                event.acceptProposedAction()
            else:
                QMessageBox.warning(self, "错误", "只能接受CSV文件！")
                event.ignore()
        except Exception:
            QMessageBox.warning(self, "错误", "处理文件时出错")
            event.ignore()

    def reset_column_widths(self):
        """重置表格列宽为默认值"""
        try:
            if hasattr(self, 'table_view'):
                # 获取窗口宽度
                window_width = self.width()
                # 计算表格总宽度
                total_width = window_width - 34  # 右侧边距设为34像素
                
                # 根据比例重新设置列宽
                self.table_view.setColumnWidth(0, int(total_width * 0.38))  # 材料名称列约38%宽度
                self.table_view.setColumnWidth(1, int(total_width * 0.12))  # 总需求约12%宽度
                self.table_view.setColumnWidth(2, int(total_width * 0.12))  # 仍需收集约12%宽度
                self.table_view.setColumnWidth(3, int(total_width * 0.12))  # 已有数量约12%宽度
                self.table_view.setColumnWidth(4, int(total_width * 0.13))  # 收集状态约13%宽度
                self.table_view.setColumnWidth(5, int(total_width * 0.13))  # 快速完成约13%宽度
                
                # 显示提示消息
                self.show_notification("已重置列宽", 2000)
        except Exception:
            self.statusBar().showMessage("重置列宽失败", 2000)
    
    # 添加显示右下角提示的方法
    def show_notification(self, message, duration=2000):
        """在窗口右下角显示提示信息"""
        try:
            # 设置提示消息
            self.save_notification.setText(message)
            
            # 调整提示框大小
            self.save_notification.adjustSize()
            
            # 计算提示框位置
            x = self.width() - self.save_notification.width() - 20
            y = self.height() - self.save_notification.height() - 20
            self.save_notification.move(x, y)
            
            # 显示提示
            self.save_notification.show()
            
            # 设置定时器隐藏提示
            if self.notification_timer.isActive():
                self.notification_timer.stop()
            self.notification_timer.start(duration)
            
            # 同时在状态栏显示信息（可选）
            self.statusBar().showMessage(message, duration)
        except Exception:
            pass
    
    def hide_notification(self):
        """隐藏提示框"""
        self.save_notification.hide()
        self.notification_timer.stop()

    def update_header_tooltips(self):
        """更新表头的 Tooltip 提示"""
        if not hasattr(self, 'table_view'):
            return
            
        header = self.table_view.horizontalHeader()
        # 只更新第5列的 Tooltip
        tooltip_text = "点击锁定/解锁此列功能" 
        if self.quick_complete_locked:
            tooltip_text += " (当前已锁定)"
        else:
            tooltip_text += " (当前已解锁)"
        header.setToolTip(tooltip_text) # 为整个表头设置提示可能不精确，需要针对section
        # 更精确的方法是设置 Section Tip (但这似乎不是标准 QHeaderView API)
        # 替代方法：可以通过样式表或事件过滤器设置，但暂时先用简单的 header tooltip
        # 或者在 handle_header_click 中更新 header 本身的 tooltip
        # 让我们在 handle_header_click 中更新 header tooltip
        pass # 逻辑移到 handle_header_click 和 initialize_table_view

    def handle_header_click(self, logicalIndex):
        """处理表头单击事件，实现降序/升序切换排序"""
        try:
            if not hasattr(self, 'table_view'):
                return

            header = self.table_view.horizontalHeader() # <-- 获取表头对象

            # 确定新的排序顺序
            if self._current_sort_column == logicalIndex:
                # 单击同一列：切换顺序
                new_order = Qt.AscendingOrder if self._current_sort_order == Qt.DescendingOrder else Qt.DescendingOrder
            else:
                # 单击新列：总是先降序
                new_order = Qt.DescendingOrder

            # 更新当前排序状态
            self._current_sort_column = logicalIndex
            self._current_sort_order = new_order

            # 应用排序 (通过代理模型，如果存在)
            if hasattr(self, 'proxy_model') and self.proxy_model is not None:
                self.proxy_model.sort(logicalIndex, new_order)
            elif hasattr(self, 'model'):
                self.model.sort(logicalIndex, new_order)

            # 更新表头排序指示器
            header.setSortIndicator(logicalIndex, new_order)

            if logicalIndex == 5:
                # 切换锁定状态
                self.quick_complete_locked = not self.quick_complete_locked
                
                if hasattr(self, 'model') and self.model:
                    # 通知模型更新状态
                    self.model.set_quick_complete_locked(self.quick_complete_locked)
                else:
                    # 如果模型不存在，直接返回或显示错误
                    return
                
                # 更新表头文本以反映状态 (模型内部会emit headerDataChanged)
                new_header_text = f"快速完成 [{'锁定' if self.quick_complete_locked else '解锁'}]"
                self.model._headers[5] = new_header_text
                # headerDataChanged 已经在 set_quick_complete_locked 中发射，这里无需重复
                # self.model.headerDataChanged.emit(Qt.Horizontal, 5, 5) 

                # 显示提示
                lock_msg = "快速完成功能已锁定" if self.quick_complete_locked else "快速完成功能已解锁"
                self.show_notification(lock_msg, 2000)

                # --- 更新 Header Tooltip --- 
                tooltip_text = f"点击锁定/解锁此列功能 (当前已{'锁定' if self.quick_complete_locked else '解锁'})"
                header.setToolTip(tooltip_text)
                # --- 结束更新 --- 

                # 如果当前是按此列排序的，清除排序指示器
                if self._current_sort_column == logicalIndex:
                    header.setSortIndicator(-1, Qt.AscendingOrder)
                    self._current_sort_column = -1 # 重置排序跟踪
                
                return # 不进行排序操作

        except Exception as e:
            QMessageBox.warning(self, "排序错误", f"排序时发生错误: {e}")

    def change_display_unit(self):
        """处理显示单位下拉框的更改"""
        try:
            new_unit = self.unit_combobox.currentText()
            if new_unit == self.current_unit:
                return

            self.current_unit = new_unit
            self.settings.setValue("display_unit", new_unit)

            if hasattr(self, 'model') and self.model:
                self.model.set_display_unit(new_unit)

                # Trigger view update using the proxy model if available
                if hasattr(self, 'proxy_model') and self.proxy_model:
                    # Invalidating the entire proxy forces a complete refresh including display roles
                    self.proxy_model.invalidate()
                else:
                    # Fallback if no proxy: signal dataChanged for relevant columns
                    num_rows = self.model.rowCount(QModelIndex()) # <-- 修正调用
                    if num_rows > 0:
                        start_col1 = self.model.index(0, 1)
                        end_col3 = self.model.index(num_rows - 1, 3)
                        # Emit for DisplayRole as that's what needs updating
                        self.model.dataChanged.emit(start_col1, end_col3, [Qt.DisplayRole])

            self.show_notification(f"显示单位已切换为: {self.current_unit}", 2000)

        except Exception as e:
            QMessageBox.warning(self, "错误", f"切换显示单位时出错: {e}")

def find_csv_files(directory):
    """查找目录中的所有CSV文件"""
    csv_files = []
    try:
        for file in os.listdir(directory):
            if file.lower().endswith('.csv'):
                csv_files.append(file)
    except Exception:
        pass
    
    return csv_files

def load_csv(file_path):
    try:
        # 尝试用不同分隔符读取
        data = pd.read_csv(file_path, sep='\t')  # 先尝试制表符分隔
        if len(data.columns) == 1:  # 如果只有一列，可能是逗号分隔
            data = pd.read_csv(file_path, sep=',')
        
        # 检查核心列数是否足够
        if len(data.columns) < 4:
            raise ValueError("CSV文件至少需要4列: 材料名称, 总需求, 仍需收集, 已有数量")
            
        # 重命名加载的前4列
        rename_map = {
            data.columns[0]: 'Item',
            data.columns[1]: 'Total',
            data.columns[2]: 'Missing',
            data.columns[3]: 'Available'
        }
        data.rename(columns=rename_map, inplace=True)
        
        # --- 确保DataFrame总是有6列 --- 
        # 模型期望的完整列名列表
        expected_columns = ['Item', 'Total', 'Missing', 'Available', '收集状态', '快速完成']
        
        # 只保留/添加必要的列
        # 我们只需要前4列有实际数据，后两列是动态生成的，但DataFrame结构需要完整
        # 确保原始4列存在且类型基本正确
        data = data[['Item', 'Total', 'Missing', 'Available']].copy() 
        # 将数值列转换为数值类型，处理可能的错误
        data['Total'] = pd.to_numeric(data['Total'], errors='coerce').fillna(0).astype(int)
        data['Missing'] = pd.to_numeric(data['Missing'], errors='coerce').fillna(0).astype(int)
        data['Available'] = pd.to_numeric(data['Available'], errors='coerce').fillna(0).astype(int)

        # 添加后两列作为占位符 (实际值由模型动态计算，但列必须存在以避免索引错误)
        data['收集状态'] = None # 或者可以预先计算，但模型会覆盖
        data['快速完成'] = None # 同上
        
        # 确保列顺序与模型headers一致
        data = data[expected_columns]
        # --- 结束确保6列 --- 
        
        return data
    except Exception as e:
        # 返回一个空的DataFrame，但包含正确的列名
        QMessageBox.critical(None, "CSV加载错误", f"加载或处理CSV文件时出错:\n{e}\n\n将创建一个空表。")
        return pd.DataFrame(columns=['Item', 'Total', 'Missing', 'Available', '收集状态', '快速完成'])

if __name__ == "__main__":
    # 设置应用和Windows风格
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("WindowsVista"))
    
    # 获取程序所在目录
    if getattr(sys, 'frozen', False):
        # 如果是打包后的应用
        app_dir = os.path.dirname(sys.executable)
    else:
        # 如果是源代码运行
        app_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 查找所有CSV文件
    csv_files = find_csv_files(app_dir)
    
    if not csv_files:
        # 没有找到CSV文件，直接创建空窗口，显示拖放提示
        window = MainWindow()
    elif len(csv_files) == 1:
        # 只有一个文件，直接加载
        csv_file = os.path.join(app_dir, csv_files[0])
        csv_data = load_csv(csv_file)
        window = MainWindow(csv_data, csv_file)
    else:
        # 有多个文件，弹出选择对话框
        dialog = FileSelectDialog(csv_files)
        if dialog.exec() == QDialog.Accepted and dialog.selected_file:
            # 用户选择了文件
            csv_file = os.path.join(app_dir, dialog.selected_file)
            csv_data = load_csv(csv_file)
            window = MainWindow(csv_data, csv_file)
        else:
            # 用户取消了选择，创建空窗口
            window = MainWindow()
    
    window.show()
    sys.exit(app.exec())