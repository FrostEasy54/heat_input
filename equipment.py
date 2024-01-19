from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QSpinBox
from PyQt6.QtWidgets import QDoubleSpinBox, QMessageBox
from PyQt6 import QtCore


class EquipmentTable:
    def AddEquipmentRow(self):
        self.EquipmentTableWidget.insertRow(
            self.EquipmentTableWidget.rowCount())
        self.EquipmentTypeComboBox()
        self.EquipmentRoomNumberSpinBox()
        self.EquipmentCountSpinBox()
        self.EquipmentPowerSpinBox()
        self.EquipmentLoadDoubleSpinBox()
        self.SetPowerForEquipmentType()
        self.MakeEquipmentCellReadOnly()
        self.CheckForSetPower()

    def RemoveEquipmentRow(self):
        if self.EquipmentTableWidget.rowCount() == 1:
            message_box = QMessageBox()
            message_box.warning(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \nЕсли в вашем помещении нет теплопоступлений от оборудования, просто укажите их количество равное 0.") # noqa E501
            message_box.setFixedSize(500, 200)
            return
        else:
            self.EquipmentTableWidget.removeRow(
                self.EquipmentTableWidget.rowCount()-1)

    # Неизменяемость клеток, в которых расчитывается Q для таблицы Оборудование
    def MakeEquipmentCellReadOnly(self):
        flags = QtCore.Qt.ItemFlag.ItemIsEnabled
        for row in range(self.EquipmentTableWidget.rowCount()):
            item = self.EquipmentTableWidget.item(row, 5)
            if item is None:
                item = QTableWidgetItem()
                self.EquipmentTableWidget.setItem(row, 5, item)
            item.setFlags(flags)

    def EquipmentRoomNumberSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 0
        current_row = self.EquipmentTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        prev_row = self.EquipmentTableWidget.rowCount() - 2
        prev_widget = self.EquipmentTableWidget.cellWidget(prev_row, col)
        prev_value = prev_widget.value() if prev_widget else 0
        sb = QSpinBox()
        sb.setMaximum(1000)
        sb.setValue(prev_value + 1)
        self.EquipmentTableWidget.setCellWidget(current_row, col, sb)

    def EquipmentTypeComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 1
        row = self.EquipmentTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Чайник", "Компьютер", "Другой"])
        self.EquipmentTableWidget.setCellWidget(row, col, cb)

    def EquipmentCountSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 4
        current_row = self.EquipmentTableWidget.rowCount() - 1
        sb = QSpinBox()
        sb.setMaximum(1000)
        self.EquipmentTableWidget.setCellWidget(current_row, col, sb)

    def EquipmentPowerSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 3
        current_row = self.EquipmentTableWidget.rowCount() - 1
        sb = QSpinBox()
        sb.setMaximum(5000)
        self.EquipmentTableWidget.setCellWidget(current_row, col, sb)

    def EquipmentLoadDoubleSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 2
        current_row = self.EquipmentTableWidget.rowCount() - 1
        sb = QDoubleSpinBox()
        sb.setMinimum(0.5)
        sb.setMaximum(0.8)
        sb.setSingleStep(0.1)
        self.EquipmentTableWidget.setCellWidget(current_row, col, sb)

    def EquipmentHeatInput(self):
        for row in range(self.EquipmentTableWidget.rowCount()):
            equipment_count_value = self.EquipmentTableWidget.cellWidget(
                row, 4).value()
            power_value = self.EquipmentTableWidget.cellWidget(row, 3).value()
            load_value = self.EquipmentTableWidget.cellWidget(row, 2).value()
            heat_input = equipment_count_value * power_value * load_value
            heat_input = float('{:.3f}'.format(heat_input))
            item = QTableWidgetItem()
            item.setData(0, f"{heat_input}")
            self.EquipmentTableWidget.setItem(row, 5, item)
            self.MakeEquipmentCellReadOnly()
        self.ActionSaveEquipment.setEnabled(True)

    def SetPowerForEquipmentType(self):
        for row in range(self.EquipmentTableWidget.rowCount()):
            equipment_type = self.EquipmentTableWidget.cellWidget(
                row, 1).currentText()
            if equipment_type == 'Чайник':
                self.EquipmentTableWidget.cellWidget(row, 3).setValue(1500)
            elif equipment_type == 'Компьютер':
                self.EquipmentTableWidget.cellWidget(row, 3).setValue(2400)
            else:
                self.EquipmentTableWidget.cellWidget(row, 3).setValue(0)

    def CheckForSetPower(self):
        for row in range(self.EquipmentTableWidget.rowCount()):
            self.EquipmentTableWidget.cellWidget(
                row, 1).activated.connect(self.SetPowerForEquipmentType)

    def ClearEquipmentTable(self):
        self.EquipmentTableWidget.setRowCount(0)
        self.AddEquipmentRow()
        self.ActionSaveEquipment.setEnabled(False)

    def ExportEquipmentTable(self, wb):
        ws_equipment = wb.create_sheet("Оборудование", 0)
        ws_equipment.merge_cells('A1:F1')
        cell_title = ws_equipment['A1']
        cell_title.value = 'Оборудование'
        column_headers = ["№ помещения", "Тип прибора", "Загрузка прибора",
                          "Мощность прибора, Вт", "Кол-во", "Q прибор, Вт"]
        for col_num, header in enumerate(column_headers, 1):
            header_cell = ws_equipment.cell(row=2, column=col_num)
            header_cell.value = header
        for row in range(self.EquipmentTableWidget.rowCount()):
            for col in range(self.EquipmentTableWidget.columnCount()):
                widget = self.EquipmentTableWidget.cellWidget(row, col)
                if widget:
                    value = widget.currentText() if isinstance(
                        widget, QComboBox) else str(widget.value())
                else:
                    item = self.EquipmentTableWidget.item(row, col)
                    if item:
                        value = item.text()
                ws_equipment.cell(row=row+3, column=col+1, value=value)

    def ImportEquipmentTable(self, wb):
        try:
            sheet_name = 'Оборудование'
            if sheet_name not in wb:
                QMessageBox.warning(
                    None, "Лист не найден", f"Лист '{sheet_name}' не найден в файле Excel.") # noqa E501
                return
            ws_equipment = wb[sheet_name]
            if not ws_equipment['A1'].value:
                QMessageBox.warning(None, "Пустой файл",
                                    "Файл не содержит данных.")
                return
            self.ClearEquipmentTable()
            for row in range(1, ws_equipment.max_row + 1):
                if row > 1:
                    self.AddEquipmentRow()
                for col, data_type in zip(range(1, 6), [int, str, float, int, int]): # noqa E501
                    cell_value = ws_equipment.cell(row=row, column=col).value
                    if not isinstance(cell_value, data_type):
                        QMessageBox.warning(
                            None, "Ошибка типа данных", f"Строка {row}, столбец {col}: Неправильный тип данных.") # noqa E501
                        return
                    widget = self.EquipmentTableWidget.cellWidget(
                        row - 1, col - 1)
                    if widget:
                        if isinstance(widget, QComboBox):
                            widget.setCurrentText(cell_value)
                        elif isinstance(widget, QSpinBox):
                            widget.setValue(int(cell_value))
                        elif isinstance(widget, QDoubleSpinBox):
                            widget.setValue(float(cell_value))
                    else:
                        item = self.EquipmentTableWidget.item(row - 1, col - 1)
                        if item:
                            item.setText(str(cell_value))
            self.tabWidget.setCurrentWidget(self.EquipmentWidget)
            QMessageBox.information(
                None, "Импорт завершен", "Данные успешно импортированы.")
        except Exception as e:
            QMessageBox.critical(
                None, "Ошибка импорта", f"Произошла ошибка при импорте данных: {str(e)}") # noqa E501
