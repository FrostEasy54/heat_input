from PyQt6.QtWidgets import QTableWidgetItem, QComboBox
from PyQt6.QtWidgets import QSpinBox, QMessageBox
from PyQt6 import QtCore


class LampsTable:
    def AddLampsRow(self):
        self.LampsTableWidget.insertRow(self.LampsTableWidget.rowCount())
        self.LampsRoomNumberSpinBox()
        self.LampsTypeComboBox()
        self.LampsPurposeComboBox()
        self.LampsCountSpinBox()
        self.MakeLampsCellReadOnly()

    def RemoveLampsRow(self):
        if self.LampsTableWidget.rowCount() == 1:
            message_box = QMessageBox()
            message_box.warning(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \nЕсли в вашем помещении нет теплопоступлений от освещения, просто укажите их количество равное 0.") # noqa E501
            message_box.setFixedSize(500, 200)
            return
        else:
            self.LampsTableWidget.removeRow(self.LampsTableWidget.rowCount()-1)

    # Неизменяемость клеток, в которых расчитывается Q для таблицы Оборудование
    def MakeLampsCellReadOnly(self):
        flags = QtCore.Qt.ItemFlag.ItemIsEnabled
        for row in range(self.LampsTableWidget.rowCount()):
            item = self.LampsTableWidget.item(row, 4)
            if item is None:
                item = QTableWidgetItem()
                self.LampsTableWidget.setItem(row, 4, item)
            item.setFlags(flags)

    def LampsRoomNumberSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 0
        current_row = self.LampsTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        prev_row = self.LampsTableWidget.rowCount() - 2
        prev_widget = self.LampsTableWidget.cellWidget(prev_row, col)
        prev_value = prev_widget.value() if prev_widget else 0
        sb = QSpinBox()
        sb.setMaximum(1000)
        sb.setValue(prev_value + 1)
        self.LampsTableWidget.setCellWidget(current_row, col, sb)

    def LampsTypeComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 1
        row = self.LampsTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Обычный", "Встроенный"])
        self.LampsTableWidget.setCellWidget(row, col, cb)

    def LampsPurposeComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 2
        row = self.LampsTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Общее освещение", "Комбинированное освещение",
                    "Декоративное освещение", "Ночник"])
        self.LampsTableWidget.setCellWidget(row, col, cb)

    def LampsCountSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 3
        current_row = self.LampsTableWidget.rowCount() - 1
        sb = QSpinBox()
        sb.setMaximum(1000)
        self.LampsTableWidget.setCellWidget(current_row, col, sb)

    def LampsHeatInput(self):
        lamp_type_dict = {'Обычный': 1, 'Встроенный': 0.4}
        lamp_purpose_dict = {'Общее освещение': 550,
                             'Комбинированное освещение': 180,
                             'Декоративное освещение': 150, 'Ночник': 25}
        for row in range(self.LampsTableWidget.rowCount()):
            lamps_count_value = self.LampsTableWidget.cellWidget(
                row, 3).value()
            coeff_value = 0.86
            type_value = lamp_type_dict.get(
                self.LampsTableWidget.cellWidget(row, 1).currentText())
            purpose_value = lamp_purpose_dict.get(
                self.LampsTableWidget.cellWidget(row, 2).currentText())
            heat_input = coeff_value * lamps_count_value * type_value * purpose_value # noqa E501
            heat_input = float('{:.3f}'.format(heat_input))
            item = QTableWidgetItem()
            item.setData(0, f"{heat_input}")
            self.LampsTableWidget.setItem(row, 4, item)
            self.MakeLampsCellReadOnly()
        self.ActionSaveLamps.setEnabled(True)

    def ClearLampsTable(self):
        self.LampsTableWidget.setRowCount(0)
        self.AddLampsRow()
        self.ActionSaveLamps.setEnabled(False)

    def ExportLampsTable(self, wb):
        ws_lamps = wb.create_sheet("Светильники", 0)
        ws_lamps.merge_cells('A1:E1')
        cell_title = ws_lamps['A1']
        cell_title.value = 'Светильники'
        column_headers = ["№ помещения", "Тип светильника",
                          "Назначение светильника", "Кол-во", "Q свет, Вт"]
        for col_num, header in enumerate(column_headers, 1):
            header_cell = ws_lamps.cell(row=2, column=col_num)
            header_cell.value = header
        for row in range(self.LampsTableWidget.rowCount()):
            for col in range(self.LampsTableWidget.columnCount()):
                widget = self.LampsTableWidget.cellWidget(row, col)
                if widget:
                    value = widget.currentText() if isinstance(
                        widget, QComboBox) else str(widget.value())
                else:
                    item = self.LampsTableWidget.item(row, col)
                    if item:
                        value = item.text()
                ws_lamps.cell(row=row+3, column=col+1, value=value)

    def ImportLampsTable(self, wb):
        try:
            sheet_name = 'Светильники'
            if sheet_name not in wb:
                QMessageBox.warning(
                    None, "Лист не найден", f"Лист '{sheet_name}' не найден в файле Excel.") # noqa E501
                return
            ws_lamps = wb[sheet_name]
            if not ws_lamps['A1'].value:
                QMessageBox.warning(None, "Пустой файл",
                                    "Файл не содержит данных.")
                return
            self.ClearLampsTable()
            for row in range(1, ws_lamps.max_row + 1):
                if row > 1:
                    self.AddLampsRow()
                for col, data_type in zip(range(1, 5), [int, str, str, int]):
                    cell_value = ws_lamps.cell(row=row, column=col).value
                    if not isinstance(cell_value, data_type):
                        QMessageBox.warning(
                            None, "Ошибка типа данных", f"Строка {row}, столбец {col}: Неправильный тип данных.") # noqa E501
                        return
                    widget = self.LampsTableWidget.cellWidget(
                        row - 1, col - 1)
                    if widget:
                        if isinstance(widget, QComboBox):
                            widget.setCurrentText(cell_value)
                        elif isinstance(widget, QSpinBox):
                            widget.setValue(int(cell_value))
                    else:
                        item = self.LampsTableWidget.item(row - 1, col - 1)
                        if item:
                            item.setText(str(cell_value))
            self.tabWidget.setCurrentWidget(self.LampsWidget)
            QMessageBox.information(
                None, "Импорт завершен", "Данные успешно импортированы.")
        except Exception as e:
            QMessageBox.critical(
                None, "Ошибка импорта", f"Произошла ошибка при импорте данных: {str(e)}") # noqa E501
