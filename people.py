from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QSpinBox, QMessageBox
from PyQt6 import QtCore


class PeopleTable():
    def AddPeopleRow(self):
        self.PeopleTableWidget.insertRow(self.PeopleTableWidget.rowCount())
        self.PeopleRoomNumberSpinBox()
        self.PeopleSexComboBox()
        self.PeopleWorkComboBox()
        self.PeopleClothesComboBox()
        self.PeopleCountSpinBox()
        self.MakePeopleCellReadOnly()

    def RemovePeopleRow(self):
        if self.PeopleTableWidget.rowCount() == 1:
            message_box = QMessageBox()
            message_box.information(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \
            \nЕсли в вашем помещении нет теплопоступлений от людей, просто укажите их количество равное 0.")
            message_box.setFixedSize(500, 200)
            return
        else:
            self.PeopleTableWidget.removeRow(
                self.PeopleTableWidget.rowCount()-1)

        # Неизменяемость клеток, в которых расчитывается Q для таблицы Люди
    def MakePeopleCellReadOnly(self):
        flags = QtCore.Qt.ItemFlag.ItemIsEnabled
        for row in range(self.PeopleTableWidget.rowCount()):
            item = self.PeopleTableWidget.item(row, 5)
            if item is None:
                item = QTableWidgetItem()
                self.PeopleTableWidget.setItem(row, 5, item)
            item.setFlags(flags)

    def PeopleRoomNumberSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 0
        current_row = self.PeopleTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        prev_row = self.PeopleTableWidget.rowCount() - 2
        prev_widget = self.PeopleTableWidget.cellWidget(prev_row, col)
        prev_value = prev_widget.value() if prev_widget else 0
        sb = QSpinBox()
        sb.setMaximum(1000)
        sb.setValue(prev_value + 1)
        self.PeopleTableWidget.setCellWidget(current_row, col, sb)

    def PeopleSexComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 1
        row = self.PeopleTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Мужской", "Женский"])
        self.PeopleTableWidget.setCellWidget(row, col, cb)

    def PeopleWorkComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 2
        row = self.PeopleTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Легкая", "Средняя", "Тяжелая"])
        self.PeopleTableWidget.setCellWidget(row, col, cb)

    def PeopleClothesComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 3
        row = self.PeopleTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки во втором столбце
        cb = QComboBox()
        cb.addItems(["Легкая", "Обычная", "Утепленная"])
        self.PeopleTableWidget.setCellWidget(row, col, cb)

    def PeopleCountSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 4
        current_row = self.PeopleTableWidget.rowCount() - 1
        sb = QSpinBox()
        sb.setMaximum(1000)
        self.PeopleTableWidget.setCellWidget(current_row, col, sb)

    def PeopleHeatInput(self):
        for row in range(self.PeopleTableWidget.rowCount()):
            people_count = self.PeopleTableWidget.cellWidget(row, 4).value()
            if self.PeopleTableWidget.cellWidget(row, 1).currentText() == 'Мужской':
                sex_coeff = 1
            else:
                sex_coeff = 0.85
            if self.PeopleTableWidget.cellWidget(row, 2).currentText() == 'Легкая':
                work_coeff = 1
            elif self.PeopleTableWidget.cellWidget(row, 2).currentText() == 'Средняя':
                work_coeff = 1.07
            else:
                work_coeff = 1.15
            if self.PeopleTableWidget.cellWidget(row, 3).currentText() == 'Легкая':
                clothes_coeff = 1
            elif self.PeopleTableWidget.cellWidget(row, 3).currentText() == 'Обычная':
                clothes_coeff = 0.65
            else:
                clothes_coeff = 0.40
            wind_speed = self.WindSpeedDoubleSpinBox.value()
            t_inside = 18
            heat_input = people_count * sex_coeff * work_coeff * clothes_coeff * \
                (2.5 + 10.36 * wind_speed ** 0.5)*(35 - t_inside)
            heat_input = float('{:.3f}'.format(heat_input))
            item = QTableWidgetItem()
            item.setData(0, f"{heat_input}")
            self.PeopleTableWidget.setItem(row, 5, item)
            self.MakePeopleCellReadOnly()

    def ClearPeopleTable(self):
        self.PeopleTableWidget.setRowCount(0)
        self.AddPeopleRow()
