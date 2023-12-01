from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QSpinBox, QMessageBox
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
            message_box.information(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \
            \nЕсли в вашем помещении нет теплопоступлений от освещения, просто укажите их количество равное 0.")
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
            heat_input = coeff_value * lamps_count_value * type_value * purpose_value
            heat_input = float('{:.3f}'.format(heat_input))
            item = QTableWidgetItem()
            item.setData(0, f"{heat_input}")
            self.LampsTableWidget.setItem(row, 4, item)
            self.MakeLampsCellReadOnly()