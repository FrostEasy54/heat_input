import os
import sys
from PyQt6.QtWidgets import QMainWindow, QHeaderView, QApplication
from PyQt6 import uic
from rooms import RoomsTable
from people import PeopleTable
from equipment import EquipmentTable
from lamps import LampsTable

ui_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'mygui.ui'))


class MyGUI(QMainWindow, RoomsTable, PeopleTable, EquipmentTable, LampsTable):

    def __init__(self):
        super(MyGUI, self).__init__()
        uic.loadUi(ui_path, self)

        # Расстягивание ширины заголовков таблиц под ширину экрана
        self.RoomsTableWidget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.PeopleTableWidget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.EquipmentTableWidget.horizontalHeader(
        ).setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.LampsTableWidget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)

        # Установка значений движения скорости ветра в помещении
        self.setWindSpeed()

        # Добавление и удаление новых рядов для таблицы Помещения
        self.AddRowRoomsPushButton.clicked.connect(self.AddRoomsRow)
        self.RemoveRowRoomsPushButton.clicked.connect(self.RemoveRoomsRow)
        # Установка виджетов для таблицы Помещения
        self.RoomTypeComboBox()
        self.RoomNumberSpinBox()
        self.RoomAreaDoubleSpinBox()
        self.MakeRoomCellReadOnly()
        # Добавление теплопоступлений с других листов на лист Помещения
        self.HeatInputPushButton.clicked.connect(self.AddPeopleHeatInput)
        self.HeatInputPushButton.clicked.connect(self.AddEquipmentHeatInput)
        self.HeatInputPushButton.clicked.connect(self.AddLampsHeatInput)
        # Сумма всех теплопоступлений
        self.HeatInputPushButton.clicked.connect(self.SumOfHeatInput)

        # Добавление и удаление новых рядов для таблицы Люди
        self.AddRowPeoplePushButton.clicked.connect(self.AddPeopleRow)
        self.RemoveRowPeoplePushButton.clicked.connect(self.RemovePeopleRow)
        # Установка виджетов для таблицы Люди
        self.PeopleSexComboBox()
        self.PeopleRoomNumberSpinBox()
        self.PeopleWorkComboBox()
        self.PeopleClothesComboBox()
        self.PeopleCountSpinBox()
        self.MakePeopleCellReadOnly()
        # Расчёт теплопоступлений от людей
        self.PeopleHeatInputPushButton.clicked.connect(self.PeopleHeatInput)

        # Добавление и удаление новых рядов для таблицы Оборудование
        self.AddRowEquipmentPushButton.clicked.connect(self.AddEquipmentRow)
        self.RemoveRowEquipmentPushButton.clicked.connect(
            self.RemoveEquipmentRow)
        # Установка виджетов для таблицы Оборудование
        self.EquipmentRoomNumberSpinBox()
        self.EquipmentTypeComboBox()
        self.EquipmentCountSpinBox()
        self.EquipmentPowerSpinBox()
        self.EquipmentLoadDoubleSpinBox()
        self.MakeEquipmentCellReadOnly()
        self.SetPowerForEquipmentType()
        self.CheckForSetPower()
        # Расчёт теплопоступлений от Оборудования на листе Оборудование
        self.EquipmentHeatInputPushButton.clicked.connect(
            self.EquipmentHeatInput)

        # Добавление и удаление новых рядов для таблицы Светильники
        self.AddRowLampsPushButton.clicked.connect(self.AddLampsRow)
        self.RemoveRowLampsPushButton.clicked.connect(self.RemoveLampsRow)
        # Установка виджетов для таблицы Светильники
        self.LampsRoomNumberSpinBox()
        self.LampsTypeComboBox()
        self.LampsPurposeComboBox()
        self.LampsCountSpinBox()
        self.MakeLampsCellReadOnly()
        # Расчёт теплопоступлений от Светильника на листе Светильники
        self.LampsHeatInputPushButton.clicked.connect(self.LampsHeatInput)

    def setWindSpeed(self):
        self.WindSpeedDoubleSpinBox.setMinimum(0.13)
        self.WindSpeedDoubleSpinBox.setMaximum(0.30)
        self.WindSpeedDoubleSpinBox.setSingleStep(0.05)


def main():
    app = QApplication(sys.argv)
    window = MyGUI()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
