from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QDoubleSpinBox
from PyQt6.QtWidgets import QSpinBox, QMessageBox
from PyQt6.QtGui import QColor
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
        self.SetRoomInnerTemp()
        self.MakeRoomCellReadOnly()
        self.MakeRoomTempCellReadOnly()

    def RemoveRoomsRow(self):
        if self.RoomsTableWidget.rowCount() == 1:
            message_box = QMessageBox()
            message_box.warning(
                None, "Единственная строка", "Вы не можете удалить единственную строку в таблице.")
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

    def MakeRoomTempCellReadOnly(self):
        col = 4
        flags = QtCore.Qt.ItemFlag.ItemIsEnabled
        for row in range(self.RoomsTableWidget.rowCount()):
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

    def SetRoomInnerTemp(self):
        col = 4
        if float(self.TempOutsideValueLabel.text()[:-2]) > -15:
            temp_inside = 18
        else:
            temp_inside = 20
        for room_row in range(self.RoomsTableWidget.rowCount()):
            temp_item = QTableWidgetItem()
            temp_item.setData(0, f'{temp_inside}')
            self.RoomsTableWidget.setItem(room_row, col, temp_item)
        self.MakeRoomTempCellReadOnly()

    def AddPeopleHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
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
                    item.setData(0, '0.0')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
                elif int_room_number_item == int_people_room_number_item:
                    heat_input_value = self.PeopleTableWidget.item(
                        people_row, 5).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.0')
                        self.RoomsTableWidget.setItem(room_row, 5, item)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Люди для помещения №{int_room_number_item} не были рассчитаны!")
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{summ_of_heat_input}')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
                elif int_room_number_item not in set(total_people_room_number):
                    item = QTableWidgetItem()
                    item.setData(0, '0.0')
                    self.RoomsTableWidget.setItem(room_row, 5, item)
        self.MakeRoomCellReadOnly()

    def AddEquipmentHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
            summ_of_heat_input = 0
            for equipment_row in range(self.EquipmentTableWidget.rowCount()):
                room_number_item = self.RoomsTableWidget.cellWidget(
                    room_row, 0)
                equipment_room_number_item = self.EquipmentTableWidget.cellWidget(
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
                if int_equipment_room_number_item not in total_equipment_room_number:
                    total_equipment_room_number.append(
                        int_equipment_room_number_item)
                if self.EquipmentTableWidget.item(equipment_row, 5) is None:
                    item = QTableWidgetItem()
                    item.setData(0, '0.0')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
                elif int_room_number_item == int_equipment_room_number_item:
                    heat_input_value = self.EquipmentTableWidget.item(
                        equipment_row, 5).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.0')
                        self.RoomsTableWidget.setItem(room_row, 6, item)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Оборудование для помещения №{int_room_number_item} не были рассчитаны!")
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{summ_of_heat_input}')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
                elif int_room_number_item not in set(total_equipment_room_number):
                    item = QTableWidgetItem()
                    item.setData(0, '0.0')
                    self.RoomsTableWidget.setItem(room_row, 6, item)
        self.MakeRoomCellReadOnly()

    def AddLampsHeatInput(self):
        for room_row in range(self.RoomsTableWidget.rowCount()):
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
                    item.setData(0, '0.0')
                    self.RoomsTableWidget.setItem(room_row, 7, item)
                elif int_room_number_item == int_lamps_room_number_item:
                    heat_input_value = self.LampsTableWidget.item(
                        lamps_row, 4).text()

                    if heat_input_value == '':
                        item = QTableWidgetItem()
                        item.setData(0, '0.0')
                        self.RoomsTableWidget.setItem(room_row, 7, item)
                        message_box = QMessageBox()
                        message_box.critical(
                            None, "Ошибка", f"Теплопоступления на листе Освещение для помещения №{int_room_number_item} не были рассчитаны!")
                        message_box.setFixedSize(500, 200)
                        break

                    float_heat_input_value = float(heat_input_value)
                    summ_of_heat_input += float_heat_input_value
                    item = QTableWidgetItem()
                    item.setData(0, f'{summ_of_heat_input}')
                    self.RoomsTableWidget.setItem(room_row, 7, item)
                elif int_room_number_item not in set(total_lamps_room_number):
                    item = QTableWidgetItem()
                    item.setData(0, '0.0')
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
                for room_col in range(5, self.RoomsTableWidget.columnCount()-1):
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
        col = 1
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.item(room_row, col) is None:
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения в строке №{room_row+1} прежде чем рассчитывать его теплопоступления!")
                message_box.setFixedSize(500, 200)
                return True
            elif self.RoomsTableWidget.item(room_row, col).text() == '':
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения в строке №{room_row+1} прежде чем рассчитывать его теплопоступления!")
                message_box.setFixedSize(500, 200)
                return True
        return False

    def RoomAreaValidation(self):
        col = 3
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.cellWidget(room_row, col) is None:
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите площадь для помещения в строке №{room_row+1} прежде чем рассчитывать его теплопоступления!")
                message_box.setFixedSize(500, 200)
                return True
            elif self.RoomsTableWidget.cellWidget(room_row, col).value() == 0:
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Площадь помещения не может быть равна 0. Пожалуйста, укажите площадь помещения в строке №{room_row+1}.")
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
        ws_rooms['A1'] = 'Город'
        ws_rooms['B1'] = self.CityComboBox.currentText()
        ws_rooms['D1'] = 't наружного воздуха'
        ws_rooms['E1'] = self.TempOutsideValueLabel.text()
        ws_rooms['G1'] = 'Период года'
        ws_rooms['H1'] = self.SeasonComboBox.currentText()
        ws_rooms['J1'] = 'V, м/с'
        ws_rooms['K1'] = self.WindSpeedDoubleSpinBox.value()
        column_headers = ["№ помещения", "Наименование помещения", "Тип помещения", "tв, °C",
                          "Площадь помещения", "Q люди, Вт", "Q приборы, Вт", "Q светильники, Вт", "Q общ, Вт"]
        for col_num, header in enumerate(column_headers, 1):
            header_cell = ws_rooms.cell(row=2, column=col_num)
            header_cell.value = header
        for row in range(self.RoomsTableWidget.rowCount()):
            for col in range(self.RoomsTableWidget.columnCount()):
                widget = self.RoomsTableWidget.cellWidget(row, col)
                if widget:
                    value = widget.currentText() if isinstance(
                        widget, QComboBox) else str(widget.value())
                else:
                    item = self.RoomsTableWidget.item(row, col)
                    if item:
                        value = item.text()
                ws_rooms.cell(row=row+3, column=col+1, value=value)

    def ImportRoomsTable(self, wb):
        try:
            sheet_name = 'Помещения'
            if sheet_name not in wb:
                QMessageBox.warning(
                    None, "Лист не найден", f"Лист '{sheet_name}' не найден в файле Excel.")
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
            self.CityComboBox.setCurrentText(city)
            self.SeasonComboBox.setCurrentText(season)
            self.WindSpeedDoubleSpinBox.setValue(float(wind_speed))
            for row in range(2, ws_rooms.max_row + 1):
                if row > 2:
                    self.AddRoomsRow()
                for col, data_type in zip(range(1, 5), [int, str, str, (int, float)]):
                    cell_value = ws_rooms.cell(row=row, column=col).value
                    if not isinstance(cell_value, data_type):
                        QMessageBox.warning(
                            None, "Ошибка типа данных", f"Строка {row}, столбец {col}: Неправильный тип данных.")
                        return
                    if col == 2:
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
                            widget.setCurrentText(cell_value)
                        elif isinstance(widget, QDoubleSpinBox):
                            widget.setValue(float(cell_value))
                        elif isinstance(widget, QSpinBox):
                            widget.setValue(int(cell_value))
            self.tabWidget.setCurrentWidget(self.RoomsWidget)
            QMessageBox.information(
                None, "Импорт завершен", "Данные успешно импортированы.")
        except Exception as e:
            QMessageBox.critical(
                None, "Ошибка импорта", f"Произошла ошибка при импорте данных: {str(e)}")

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
            q_lighting_item = self.RoomsTableWidget.item(room_row, 6)
            q_equipment_item = self.RoomsTableWidget.item(room_row, 7)
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
