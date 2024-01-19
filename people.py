from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QSpinBox
from PyQt6.QtWidgets import QMessageBox
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
            message_box.warning(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \nЕсли в вашем помещении нет теплопоступлений от людей, просто укажите их количество равное 0.")  # noqa E501
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
        wind_speed = self.WindSpeedDoubleSpinBox.value()
        room_number_list = []
        for row in range(self.PeopleTableWidget.rowCount()):
            people_room_number = self.PeopleTableWidget.cellWidget(
                row, 0).value()
            for room_row in range(self.RoomsTableWidget.rowCount()):
                room_number_list.append(self.RoomsTableWidget.cellWidget(
                    room_row, 0).value())
            if people_room_number in room_number_list:
                t_inside = float(self.RoomsTableWidget.cellWidget(
                    room_number_list.index(people_room_number), 4).value())
            else:
                self.tabWidget.setCurrentWidget(self.RoomsWidget)
                QMessageBox.warning(
                    None, "Помещение отсутствует", f"Невозможно узнать значение температуры в помещении, так как \nПомещение с номером {people_room_number} на листе Помещения отсутcтвует!")  # noqa E501
                return
            people_count = self.PeopleTableWidget.cellWidget(row, 4).value()
            if self.PeopleTableWidget.cellWidget(row, 1).currentText() == 'Мужской':  # noqa E501
                sex_coeff = 1
            else:
                sex_coeff = 0.85
            if self.PeopleTableWidget.cellWidget(row, 2).currentText() == 'Легкая':  # noqa E501
                work_coeff = 1
            elif self.PeopleTableWidget.cellWidget(row, 2).currentText() == 'Средняя':  # noqa E501
                work_coeff = 1.07
            else:
                work_coeff = 1.15
            if self.PeopleTableWidget.cellWidget(row, 3).currentText() == 'Легкая':  # noqa E501
                clothes_coeff = 1
            elif self.PeopleTableWidget.cellWidget(row, 3).currentText() == 'Обычная':  # noqa E501
                clothes_coeff = 0.65
            else:
                clothes_coeff = 0.40
            heat_input = people_count * sex_coeff * work_coeff * \
                clothes_coeff * (2.5 + 10.36 * wind_speed **
                                 0.5)*(35 - t_inside)
            heat_input = float('{:.3f}'.format(heat_input))
            item = QTableWidgetItem()
            item.setData(0, f"{heat_input}")
            self.PeopleTableWidget.setItem(row, 5, item)
            self.MakePeopleCellReadOnly()
        self.ActionSavePeople.setEnabled(True)

    def ClearPeopleTable(self):
        self.PeopleTableWidget.setRowCount(0)
        self.AddPeopleRow()
        self.ActionSavePeople.setEnabled(False)

    def ExportPeopleTable(self, wb):
        ws_people = wb.create_sheet("Люди", 0)
        for row in range(self.PeopleTableWidget.rowCount()):
            for col in range(self.PeopleTableWidget.columnCount()):
                widget = self.PeopleTableWidget.cellWidget(row, col)
                if widget:
                    if col in [0, 4]:
                        value = int(widget.value())
                    elif col in [1, 2, 3]:
                        value = str(widget.currentText())
                else:
                    item = self.PeopleTableWidget.item(row, col)
                    if item:
                        value = float(item.text())
                ws_people.cell(row=row+1, column=col+1, value=value)

    def ImportPeopleTable(self, wb):
        try:
            sheet_name = 'Люди'
            if sheet_name not in wb:
                QMessageBox.warning(
                    None, "Лист не найден", f"Лист '{sheet_name}' не найден в файле Excel.")  # noqa E501
                return
            ws_people = wb[sheet_name]
            if not ws_people['A1'].value:
                QMessageBox.warning(None, "Пустой файл",
                                    "Файл не содержит данных.")
                return
            self.ClearPeopleTable()
            for row in range(1, ws_people.max_row + 1):
                if row > 1:
                    self.AddPeopleRow()
                for col, data_type in zip(range(1, 7), [int, str, str, str, int, (int, float)]):  # noqa E501
                    cell_value = ws_people.cell(row=row, column=col).value
                    if not isinstance(cell_value, data_type):
                        QMessageBox.warning(
                            None, "Ошибка типа данных", f"Строка {row}, столбец {col}: Неправильный тип данных.")  # noqa E501
                        return
                    widget = self.PeopleTableWidget.cellWidget(
                        row - 1, col - 1)
                    if widget:
                        if isinstance(widget, QComboBox):
                            widget.setCurrentText(cell_value)
                        elif isinstance(widget, QSpinBox):
                            widget.setValue(int(cell_value))
                    else:
                        item = self.PeopleTableWidget.item(row - 1, col - 1)
                        if item:
                            item.setText(str(cell_value))
            self.tabWidget.setCurrentWidget(self.PeopleWidget)
            QMessageBox.information(
                None, "Импорт завершен", "Данные успешно импортированы.")
        except Exception as e:
            QMessageBox.critical(
                None, "Ошибка импорта", f"Произошла ошибка при импорте данных: {str(e)}")  # noqa E501
