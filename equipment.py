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
            message_box.information(None, "Единственная строка", "Вы не можете удалить единственную строку в таблице. \
            \nЕсли в вашем помещении нет теплопоступлений от оборудования, просто укажите их количество равное 0.")
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
