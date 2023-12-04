import os
import sys
import openpyxl
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
        # Установка списка городов и температуры в таблицу
        self.GetClimateCity()
        self.UpdateTempLabel(0)
        # Установка значений движения скорости ветра в помещении
        self.SetWindSpeed()
        self.SeasonComboBox.currentTextChanged.connect(self.SetWindSpeed)

        # Соединение кнопок очистки
        self.actionClearRooms.triggered.connect(self.ClearRoomsTable)
        self.actionClearPeople.triggered.connect(self.ClearPeopleTable)
        self.actionClearEquipment.triggered.connect(self.ClearEquipmentTable)
        self.actionClearLamps.triggered.connect(self.ClearLampsTable)
        self.actionClearAll.triggered.connect(self.ClearAllTables)

        # Добавление и удаление новых рядов для таблицы Помещения
        self.AddRowRoomsPushButton.clicked.connect(self.AddRoomsRow)
        self.RemoveRowRoomsPushButton.clicked.connect(self.RemoveRoomsRow)
        # Установка виджетов для таблицы Помещения
        self.RoomTypeComboBox()
        self.RoomNumberSpinBox()
        self.RoomAreaDoubleSpinBox()
        self.SetRoomInnerTemp()
        self.CityComboBox.currentIndexChanged.connect(self.SetRoomInnerTemp)
        self.MakeRoomCellReadOnly()
        self.MakeRoomTempCellReadOnly()
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

    def GetClimateCity(self):
        path = 'climate_db.xlsx'
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        row = 651
        self.city_temp_db = {}
        for city, temp in zip(ws['A'][9:row], ws['L'][9:row]):
            self.city_temp_db[city.value] = temp.value
        self.CityComboBox.addItems(self.city_temp_db.keys())
        self.CityComboBox.currentIndexChanged.connect(self.UpdateTempLabel)

    def UpdateTempLabel(self, index):
        current_city = self.CityComboBox.currentText()
        temp = self.city_temp_db.get(current_city, 'Н/Д')
        self.TempOutsideValueLabel.setText(f'{temp}°C')

    def SetWindSpeed(self):
        if self.SeasonComboBox.currentText() == "Теплый":
            self.WindSpeedDoubleSpinBox.setMinimum(0.1)
            self.WindSpeedDoubleSpinBox.setMaximum(0.6)
        else:
            self.WindSpeedDoubleSpinBox.setMinimum(0.0)
            self.WindSpeedDoubleSpinBox.setMaximum(0.5)
        self.WindSpeedDoubleSpinBox.setSingleStep(0.1)

    def ClearAllTables(self):
        self.ClearPeopleTable()
        self.ClearRoomsTable()
        self.ClearEquipmentTable()
        self.ClearLampsTable()


def main():
    app = QApplication(sys.argv)
    window = MyGUI()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
