from PyQt6.QtWidgets import QTableWidgetItem, QComboBox, QDoubleSpinBox
from PyQt6.QtWidgets import QSpinBox, QMessageBox
from PyQt6 import QtCore
from PyQt6.QtWidgets import *

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

    def RemoveRoomsRow(self):
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
        col = 4
        current_row = self.RoomsTableWidget.rowCount() - 1
        # Получаем значение номера помещения из предыдущей строки
        sb = QDoubleSpinBox()
        sb.setMaximum(1000)
        self.RoomsTableWidget.setCellWidget(current_row, col, sb)

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
        if self.RoomNameValidation():
            return

        for room_row in range(self.RoomsTableWidget.rowCount()):
            summ_of_heat_input = 0
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
            summ_of_heat_input = float('{:.3f}'.format(summ_of_heat_input))
            item = QTableWidgetItem()
            item.setData(0, f'{summ_of_heat_input}')
            self.RoomsTableWidget.setItem(room_row, 8, item)
        self.MakeRoomCellReadOnly()
        print(total_room_number)

    def RoomNameValidation(self):
        col = 1
        for room_row in range(self.RoomsTableWidget.rowCount()):
            if self.RoomsTableWidget.item(room_row, col) is None:
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения №{room_row+1} прежде чем рассчитывать его теплопоступления!")
                message_box.setFixedSize(500, 200)
                return True
            elif self.RoomsTableWidget.item(room_row, col).text() == '':
                message_box = QMessageBox()
                message_box.critical(
                    None, "Ошибка", f"Пожалуйста, укажите название для помещения №{room_row+1} прежде чем рассчитывать его теплопоступления!")
                message_box.setFixedSize(500, 200)
                return True
        return False
