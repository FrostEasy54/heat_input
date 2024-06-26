from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QDoubleSpinBox
from PyQt6.QtWidgets import QSpinBox, QMessageBox
from PyQt6 import QtCore
import matplotlib.pyplot as plt
import numpy as np

total_room_number = []
total_people_room_number = []
total_equipment_room_number = []
total_lamps_room_number = []


class RoomsTable():
    def AddRoomsRow(self):
        self.RoomsTableWidget.insertRow(self.RoomsTableWidget.rowCount())
        self.RoomTypeComboBox()
        self.RoomNumberSpinBox()
        self.RoomAreaDoubleSpinBox()
        self.MakeRoomCellReadOnly()
        self.RoomTempDoubleSpinBox()

    def RemoveRoomsRow(self):
        if self.RoomsTableWidget.rowCount() == 1:
            message_box = QMessageBox()
            message_box.warning(
                None, "Единственная строка", "Вы не можете удалить единственную строку в таблице.")  # noqa E501
            message_box.setFixedSize(500, 200)
            return
        else:
            self.RoomsTableWidget.removeRow(self.RoomsTableWidget.rowCount()-1)

    # Неизменяемость клеток, в которых расчитывается Q для таблицы Помещения
    def MakeRoomCellReadOnly(self):
        flags = QtCore.Qt.ItemFlag.ItemIsEnabled
        for row in range(self.RoomsTableWidget.rowCount()):
            for col in range(5, self.RoomsTableWidget.columnCount()):
                item = self.RoomsTableWidget.item(row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.RoomsTableWidget.setItem(row, col, item)
                item.setFlags(flags)

    def RoomTypeComboBox(self):
        # Указываем столбец, для которого нужно установить combobox
        col = 2
        row = self.RoomsTableWidget.rowCount() - 1
        # Устанавливаем combobox для каждой ячейки в третьем столбце
        cb = QComboBox()
        cb.addItems(["Производственный", "Бытовой"])
        self.RoomsTableWidget.setCellWidget(row, col, cb)
        if cb:
            cb.currentTextChanged.connect(self.UpdateRoomInnerTempMinMax)

    def RoomNumberSpinBox(self):
        # Указываем столбец, для которого нужно установить SpinBox
        col = 0
        current_row = self.RoomsTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        prev_row = self.RoomsTableWidget.rowCount() - 2
        prev_widget = self.RoomsTableWidget.cellWidget(prev_row, col)
        prev_value = prev_widget.value() if prev_widget else 0
        sb = QSpinBox()
        sb.setMaximum(1000)
        sb.setValue(prev_value + 1)
        self.RoomsTableWidget.setCellWidget(current_row, col, sb)

    def RoomAreaDoubleSpinBox(self):
        col = 3
        current_row = self.RoomsTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        sb = QDoubleSpinBox()
        sb.setMaximum(1000)
        self.RoomsTableWidget.setCellWidget(current_row, col, sb)

    def RoomTempDoubleSpinBox(self):
        col = 4
        current_row = self.RoomsTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        sb = QDoubleSpinBox()
        sb.setMinimum(-50)
        sb.setMaximum(50)
        self.RoomsTableWidget.setCellWidget(current_row, col, sb)

    def UpdateRoomInnerTempMinMax(self):
        col = 4
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, 2).currentText() == 'Бытовой':  # noqa E501
                if self.SeasonComboBox.currentText() == 'Холодный':
                    self.RoomsTableWidget.cellWidget(
                        room_row, col).setMinimum(18)
                    self.RoomsTableWidget.cellWidget(
                        room_row, col).setMaximum(24)
                elif self.SeasonComboBox.currentText() == 'Теплый':
                    self.RoomsTableWidget.cellWidget(
                        room_row, col).setMinimum(20)
                    self.RoomsTableWidget.cellWidget(
                        room_row, col).setMaximum(28)
            else:
                self.RoomsTableWidget.cellWidget(room_row, col).setMinimum(-50)
                self.RoomsTableWidget.cellWidget(room_row, col).setMaximum(50)

    def AddPeopleHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, 2).currentText() == 'Бытовой':  # noqa E501
                continue
            summ_of_heat_input = 0
            for people_row in range(self.PeopleTableWidget.rowCount()):
                room_number_item = self.RoomsTableWidget.cellWidget(
                    room_row, 0)
                people_room_number_item = self.PeopleTableWidget.cellWidget(
                    people_row, 0)
                if people_room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Люди!")
                    message_box.setFixedSize(500, 200)
                    break
                elif room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Помещения!")
                    message_box.setFixedSize(500, 200)
                    break
                int_room_number_item = int(room_number_item.value())
                if int_room_number_item not in total_room_number:
                    total_room_number.append(int_room_number_item)
                int_people_room_number_item = int(
                    people_room_number_item.value())
                if int_people_room_number_item not in total_people_room_number:
                    total_people_room_number.append(
                        int_people_room_number_item)
                if self.PeopleTableWidget.item(people_row, 5) is None:
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
                elif int_room_number_item == int_people_room_number_item:
                    heat_input_value = self.PeopleTableWidget.item(
                        people_row, 5).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.000')
                        self.RoomsTableWidget.setItem(room_row, 5, item)
                        self.tabWidget.setCurrentWidget(self.PeopleWidget)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Люди для помещения №{int_room_number_item} не были рассчитаны!")  # noqa E501
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{round(summ_of_heat_input, 3)}')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
                elif int_room_number_item not in set(total_people_room_number):
                    self.tabWidget.setCurrentWidget(self.PeopleWidget)
                    QMessageBox().critical(None, "Ошибка",
                                           f"Помещение №{int_room_number_item} отсутствует в списке помещений на листе Люди!")  # noqa E501
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
        self.MakeRoomCellReadOnly()

    def AddEquipmentHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, 2).currentText() == 'Бытовой':  # noqa E501
                continue
            summ_of_heat_input = 0
            for equipment_row in range(self.EquipmentTableWidget.rowCount()):
                room_number_item = self.RoomsTableWidget.cellWidget(
                    room_row, 0)
                equipment_room_number_item = self.EquipmentTableWidget.cellWidget(  # noqa E501
                    equipment_row, 0)
                if equipment_room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Оборудование!")
                    message_box.setFixedSize(500, 200)
                    break
                elif room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Помещения!")
                    message_box.setFixedSize(500, 200)
                    break
                int_room_number_item = int(room_number_item.value())
                if int_room_number_item not in total_room_number:
                    total_room_number.append(int_room_number_item)
                int_equipment_room_number_item = int(
                    equipment_room_number_item.value())
                if int_equipment_room_number_item not in total_equipment_room_number:  # noqa E501
                    total_equipment_room_number.append(
                        int_equipment_room_number_item)
                if self.EquipmentTableWidget.item(equipment_row, 5) is None:
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
                elif int_room_number_item == int_equipment_room_number_item:
                    heat_input_value = self.EquipmentTableWidget.item(
                        equipment_row, 5).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.000')
                        self.RoomsTableWidget.setItem(room_row, 6, item)
                        self.tabWidget.setCurrentWidget(self.EquipmentWidget)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Оборудование для помещения №{int_room_number_item} не были рассчитаны!")  # noqa E501
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{round(summ_of_heat_input, 3)}')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
                elif int_room_number_item not in set(total_equipment_room_number):  # noqa E501
                    self.tabWidget.setCurrentWidget(self.EquipmentWidget)
                    QMessageBox().critical(None, "Ошибка",
                                           f"Помещение №{int_room_number_item} отсутствует в списке помещений на листе Оборудование!")  # noqa E501
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
        self.MakeRoomCellReadOnly()

    def AddLampsHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, 2).currentText() == 'Бытовой':  # noqa E501
                continue
            summ_of_heat_input = 0
            for lamps_row in range(self.LampsTableWidget.rowCount()):
                room_number_item = self.RoomsTableWidget.cellWidget(
                    room_row, 0)
                lamps_room_number_item = self.LampsTableWidget.cellWidget(
                    lamps_row, 0)
                if lamps_room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Освещение!")
                    message_box.setFixedSize(500, 200)
                    break
                elif room_number_item is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Вы не указали помещения \
                             на листе Помещения!")
                    message_box.setFixedSize(500, 200)
                    break
                int_room_number_item = int(room_number_item.value())
                if int_room_number_item not in total_room_number:
                    total_room_number.append(int_room_number_item)
                int_lamps_room_number_item = int(
                    lamps_room_number_item.value())
                if int_lamps_room_number_item not in total_lamps_room_number:
                    total_lamps_room_number.append(int_lamps_room_number_item)
                if self.LampsTableWidget.item(lamps_row, 4) is None:
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 7, item)
                elif int_room_number_item == int_lamps_room_number_item:
                    heat_input_value = self.LampsTableWidget.item(
                        lamps_row, 4).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.000')
                        self.RoomsTableWidget.setItem(room_row, 7, item)
                        self.tabWidget.setCurrentWidget(self.LampsWidget)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Освещение для помещения №{int_room_number_item} не были рассчитаны!")  # noqa E501
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{round(summ_of_heat_input, 3)}')
                    self.RoomsTableWidget.setItem(room_row, 7, item)
                elif int_room_number_item not in set(total_lamps_room_number):
                    self.tabWidget.setCurrentWidget(self.LampsWidget)
                    QMessageBox().critical(None, "Ошибка",
                                           f"Помещение №{int_room_number_item} отсутствует в списке помещений на листе Освещение!")  # noqa E501
                    item = QTableWidgetItem()
                    item.setData(0, '0.000')
                    self.RoomsTableWidget.setItem(room_row, 7, item)
        self.MakeRoomCellReadOnly()

    def SumOfHeatInput(self):
        if self.RoomNameValidation() or self.RoomAreaValidation():
            return
        for room_row in range(self.RoomsTableWidget.rowCount()):
            summ_of_heat_input = 0
            room_type_item = self.RoomsTableWidget.cellWidget(room_row, 2)
            room_area_item = self.RoomsTableWidget.cellWidget(room_row, 3)
            if room_type_item.currentText() == "Производственный":
                # Выполняем расчеты для производственных помещений
                for room_col in range(5, self.RoomsTableWidget.columnCount()-1):  # noqa E501
                    summ_of_heat_input += float(
                        self.RoomsTableWidget.item(room_row, room_col).text())
                if summ_of_heat_input == '':
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", "Теплопоступления для помещения \
                            не были рассчитаны!")
                    message_box.setFixedSize(500, 200)
                    break
            elif room_type_item.currentText() == "Бытовой":
                # Умножаем площадь на коэффициент q = 21
                q_1 = 21
                room_area = room_area_item.value()
                summ_of_heat_input = room_area * q_1
                for col in range(5, self.RoomsTableWidget.columnCount()-1):
                    item = self.RoomsTableWidget.item(room_row, col)
                    if item:
                        item.setText('')
            # Устанавливаем результат в соответствующую ячейку
            item = QTableWidgetItem()
            item.setData(0, f'{summ_of_heat_input:.3f}')
            self.RoomsTableWidget.setItem(room_row, 8, item)

        self.MakeRoomCellReadOnly()
        self.ActionSaveRooms.setEnabled(True)
        self.ActionSaveAll.setEnabled(True)
        self.actionHeatInputGraph.setEnabled(True)

    def RoomNameValidation(self):
        self.tabWidget.setCurrentWidget(self.RoomsWidget)
        col = 1
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.item(room_row, col) is None:
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения в строке №{room_row+1}, прежде чем рассчитывать его теплопоступления!")  # noqa E501
                message_box.setFixedSize(500, 200)
                return True
            elif self.RoomsTableWidget.item(room_row, col).text() == '':
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения в строке №{room_row+1}, прежде чем рассчитывать его теплопоступления!")  # noqa E501
                message_box.setFixedSize(500, 200)
                return True
        return False

    def RoomAreaValidation(self):
        self.tabWidget.setCurrentWidget(self.RoomsWidget)
        col = 3
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, 2).currentText() == "Бытовой":  # noqa E501
                if self.RoomsTableWidget.cellWidget(room_row, col) is None:
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", f"Пожалуйста, укажите площадь для помещения в строке №{room_row+1}, прежде чем рассчитывать его теплопоступления!")  # noqa E501
                    message_box.setFixedSize(500, 200)
                    return True
                elif self.RoomsTableWidget.cellWidget(room_row, col).value() == 0:  # noqa E501
                    message_box = QMessageBox()
                    message_box.critical(
                        None, "Ошибка", f"Площадь помещения не может быть равна 0. Пожалуйста, укажите площадь помещения в строке №{room_row+1}.")  # noqa E501
                    message_box.setFixedSize(500, 200)
                    return True
        return False

    def ClearRoomsTable(self):
        self.RoomsTableWidget.setRowCount(0)
        self.AddRoomsRow()
        self.CityComboBox.setCurrentIndex(0)
        self.SeasonComboBox.setCurrentIndex(0)
        self.WindSpeedDoubleSpinBox.setValue(0.1)
        self.ActionSaveRooms.setEnabled(False)

    def ExportRoomsTable(self, wb):
        ws_rooms = wb.create_sheet("Помещения", 0)
        ws_rooms['A1'] = str(self.CityComboBox.currentText())
        ws_rooms['B1'] = str(self.SeasonComboBox.currentText())
        ws_rooms['C1'] = float(self.WindSpeedDoubleSpinBox.value())
        for row in range(self.RoomsTableWidget.rowCount()):
            for col in range(self.RoomsTableWidget.columnCount()):
                widget = self.RoomsTableWidget.cellWidget(row, col)
                if widget:
                    if col == 0:
                        value = int(widget.value())
                    elif col == 2:
                        value = str(widget.currentText())
                    elif col in [3, 4]:
                        value = float(widget.value())
                    else:
                        value = str(widget.value())
                else:
                    item = self.RoomsTableWidget.item(row, col)
                    if item:
                        if col in [5, 6, 7, 8]:
                            value = float(item.text())
                        else:
                            value = str(item.text())
                ws_rooms.cell(row=row+2, column=col+1, value=value)

    def ImportRoomsTable(self, wb):
        try:
            sheet_name = 'Помещения'
            if sheet_name not in wb:
                QMessageBox.warning(
                    None, "Лист не найден", f"Лист '{sheet_name}' не найден в файле Excel.") # noqa E501
                return
            ws_rooms = wb[sheet_name]
            if not ws_rooms['A1'].value or not ws_rooms['A2'].value:
                QMessageBox.warning(None, "Пустой файл",
                                    "Файл не содержит данных.")
                return
            self.ClearRoomsTable()
            city = ws_rooms['A1'].value
            season = ws_rooms['B1'].value
            wind_speed = ws_rooms['C1'].value
            self.CityComboBox.setCurrentText(str(city))
            self.SeasonComboBox.setCurrentText(str(season))
            self.WindSpeedDoubleSpinBox.setValue(float(wind_speed))
            for row in range(2, ws_rooms.max_row + 1):
                if row > 2:
                    self.AddRoomsRow()
                for col, data_type in zip(range(1, 10), [int, str, str, (int, float), (int, float), (int, float), (int, float), (int, float), (int, float)]): # noqa E501
                    cell_value = ws_rooms.cell(row=row, column=col).value
                    if not isinstance(cell_value, data_type):
                        QMessageBox.warning(
                            None, "Ошибка типа данных", f"Строка {row}, столбец {col}: Неправильный тип данных.") # noqa E501
                        return
                    if col in [2, 6, 7, 8, 9]:
                        item = self.RoomsTableWidget.item(row - 2, col - 1)
                        if not item:
                            item = QTableWidgetItem()
                            self.RoomsTableWidget.setItem(
                                row - 2, col - 1, item)
                        item.setText(str(cell_value))
                    else:
                        widget = self.RoomsTableWidget.cellWidget(
                            row - 2, col - 1)
                        if not widget:
                            widget = QComboBox()
                            self.RoomsTableWidget.setCellWidget(
                                row - 2, col - 1, widget)
                        if isinstance(widget, QComboBox):
                            widget.setCurrentText(str(cell_value))
                        elif isinstance(widget, QDoubleSpinBox):
                            widget.setValue(float(cell_value))
                        elif isinstance(widget, QSpinBox):
                            widget.setValue(int(cell_value))
            self.tabWidget.setCurrentWidget(self.RoomsWidget)
            QMessageBox.information(
                None, "Импорт завершен", "Данные успешно импортированы.")
        except Exception as e:
            QMessageBox.critical(
                None, "Ошибка импорта", f"Произошла ошибка при импорте данных: {str(e)}") # noqa E501

    def showHistogram(self):
        # Создаем новое окно для гистограммы
        plt.figure(figsize=(8, 6))
        plt.title("Гистограмма теплопоступлений")

        # Задаем данные для гистограммы
        sources = ['Q люди, Вт', 'Q освещение, Вт', 'Q приборы, Вт']
        values = []
        q_people_values = 0
        q_lighting_values = 0
        q_equipment_values = 0

        for room_row in range(self.RoomsTableWidget.rowCount()):
            q_people_item = self.RoomsTableWidget.item(room_row, 5)
            q_equipment_item = self.RoomsTableWidget.item(room_row, 6)
            q_lighting_item = self.RoomsTableWidget.item(room_row, 7)
            if q_people_item and q_people_item.text() != '':
                q_people_values += float(q_people_item.text())
            if q_lighting_item and q_lighting_item.text() != '':
                q_lighting_values += float(q_lighting_item.text())
            if q_equipment_item and q_equipment_item.text() != '':
                q_equipment_values += float(q_equipment_item.text())
        values.append(q_people_values)
        values.append(q_lighting_values)
        values.append(q_equipment_values)

        # Создаем объект Bar для гистограммы
        bar_width = 0.2
        indices = np.arange(len(sources))

        bars = plt.bar(indices, values, width=bar_width, color=['r', 'g', 'b'])

        # Добавляем точные значения над каждым столбцом
        for bar, value in zip(bars, values):
            plt.text(bar.get_x() + bar.get_width() / 2,
                     bar.get_height() + 0.1, f'{value:.2f}', ha='center')

        # Добавляем метки и заголовок
        plt.xlabel('\nИсточники теплопоступления')
        plt.ylabel('Суммарное теплопоступление')
        plt.xticks(indices, sources)

        # Отображаем гистограмму
        plt.show()
