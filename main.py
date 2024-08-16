import sys
import os
import glob
import json
from calendar import monthrange
import xlwings as xw
from xlwings.constants import FindLookIn, LookAt, SearchDirection, SearchOrder
from PyQt6 import uic
from PyQt6.QtCore import QSettings, QStringListModel, QAbstractListModel, QModelIndex, Qt, QDateTime, QTime, \
    QItemSelectionModel, QDate, QSignalBlocker
from PyQt6.QtWidgets import (QMainWindow, QDialog ,QPushButton, QApplication, QTimeEdit,
                             QMessageBox, QLineEdit, QLabel, QComboBox, QDateTimeEdit,
                             QCheckBox, QFileDialog, QSpinBox, QFileDialog,
                             QRadioButton, QGroupBox, QListView)

# Structure constants
PLAN_DAYTYPE_COL = 'D'
PLAN_ABSENCE_COL = 'B'
PLAN_DAYOFMONTH_COL = 'A'
PLAN_WEEKDAY_COL = 'C'
PLAN_STARTING_ROW = 13
WORKTIME_TYPE_COL = 'C'
WORKTIME_START_DAY_COL = 'D'
WORKTIME_START_TIME_COL = 'E'
WORKTIME_END_DAY_COL = 'F'
WORKTIME_END_TIME_COL = 'G'
WORKTIME_COMMENTS_COL = 'J'
WORKTIME_STARTING_ROW = 10
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
WORKTYPES = ["Office Hours", "Remote Work", "Overtime (paid)", "Overtime (time compensated)"]
ACTIONS = ["Working usual times", "Working custom times", "Vacation", "Half Day Vacation", "Sick", "Shift Compensation", "Flexible Time Compensation"]


class WeekdayUsualsList(QAbstractListModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._work_times = dict()  # 0 = Monday, 1 = Tuesday, etc
        self._weekday = None

    def set_weekday(self, weekday):
        self.beginResetModel()
        self._weekday = str(weekday)
        if self._weekday not in self._work_times:
            print(f"{self._weekday} is not a there{self._work_times}")
            self._work_times[self._weekday] = list()
        self.endResetModel()

    def __iter__(self):
        return iter(self._work_times[self._weekday])

    def __getitem__(self, i):
        return self._work_times[self._weekday][i]

    def __len__(self):
        if self._weekday is None:
            return 0
        return len(self._work_times[self._weekday])

    def rowCount(self, parent=None):
        if self._weekday is None:
            return 0
        return len(self._work_times[self._weekday])

    def data(self, index, role):
        data = self._work_times[self._weekday][index.row()]
        if role == Qt.ItemDataRole.ToolTipRole or role == Qt.ItemDataRole.DisplayRole:
            return f"{data["start"].toString('HH:mm')} - {data["end"].toString('HH:mm')} {WORKTYPES[data['type']]}"
        elif role == Qt.ItemDataRole.UserRole:
            return data

    def removeRow(self, index):
        self.beginRemoveRows(QModelIndex(), index, index)
        del self._work_times[self._weekday][index]
        self.endRemoveRows()

    def find(self, day_of_week_index):
        return self._work_times[str(day_of_week_index)]

    def add_work_time(self, work_time):
        index = QModelIndex()
        self.beginInsertRows(index, self.rowCount(index), self.rowCount(index))
        self._work_times[self._weekday].append(work_time)
        self._work_times[self._weekday].sort(key=lambda x: x["start"], reverse=False)
        self.endInsertRows()

    def modify_work_time(self, index, work_time):
        self.beginResetModel()
        self._work_times[self._weekday][index.row()] = work_time
        self._work_times[self._weekday].sort(key=lambda x: x["start"], reverse=False)
        self.endResetModel()

    def getUsuals(self):
        print("getUsuals")
        transformed_data = {}
        for key, value_list in self._work_times.items():
            transformed_value_list = []
            for time_dict in value_list:
                transformed_time_dict = {
                    'start': {'hour': time_dict['start'].hour(), 'min': time_dict['start'].minute()},
                    'end': {'hour': time_dict['end'].hour(), 'min': time_dict['end'].minute()},
                    'type': time_dict['type']
                }
                transformed_value_list.append(transformed_time_dict)
            transformed_data[key] = transformed_value_list
        return transformed_data

    def setUsuals(self, data):
        self.beginResetModel()
        self._work_times = {}
        print("setUsuals")
        for key, value_list in data.items():
            original_value_list = []
            for time_dict in value_list:
                original_time_dict = {
                    'start': QTime(time_dict['start']['hour'], time_dict['start']['min']),
                    'end': QTime(time_dict['end']['hour'], time_dict['end']['min']),
                    'type': time_dict['type']
                }
                original_value_list.append(original_time_dict)
            original_value_list.sort(key=lambda x: x["start"], reverse=False)
            self._work_times[key] = original_value_list
        self.endResetModel()


class WorktimeListModel(QAbstractListModel):
    def __init__(self, work_times=None, parent=None):
        super().__init__(parent)
        if work_times is None:
            self._work_times = list()
        else:
            self._work_times = [{
                'start': QTime(x['start']['h'], x['start']['m']),
                'end': QTime(x['end']['h'], x['end']['m']),
                'type': x['type']
            } for x in work_times]

    def __iter__(self):
        return iter(self._work_times)

    def __getitem__(self, i):
        return self._work_times[i]

    def __len__(self):
        return len(self._work_times)

    def rowCount(self, parent=None):
        return len(self._work_times)

    def data(self, index, role):
        if 0 <= index.row() < len(self._work_times):
            #print(index.row(), len(self._work_times))
            data = self._work_times[index.row()]
            if role == Qt.ItemDataRole.ToolTipRole or role == Qt.ItemDataRole.DisplayRole:
                return f"{data["start"].toString('HH:mm')} - {data["end"].toString('HH:mm')}: {WORKTYPES[data["type"]]}"
            elif role == Qt.ItemDataRole.UserRole:
                return data
        else:
            print("Something went wrong")

    # def resetData(self, new_data=None):
    #     self.beginResetModel()
    #     if new_data is not None:
    #         self._work_times = list()
    #     else:
    #         self._work_times = new_data
    #     self.endResetModel()

    def removeRow(self, index):
        self.beginRemoveRows(QModelIndex(), index, index)
        del self._work_times[index]
        self.endRemoveRows()

    def getData(self):
        items = [{
                    'start': dict({'h': x['start'].hour(), 'm': x['start'].minute()}),
                    'end': dict({'h': x['end'].hour(), 'm': x['end'].minute()}),
                    'type': x['type']
                  } for x in self._work_times]
        return items

    def getWorkTimes(self):
        return self._work_times

    def addItem(self, data):
        index = QModelIndex()
        self.beginInsertRows(index, self.rowCount(index), self.rowCount(index))
        self._work_times.append(data)
        self._work_times.sort(key=lambda x: x["start"], reverse=False)
        self.endInsertRows()

    def modifyItem(self, index, data):
        self.beginResetModel()
        self._work_times[index.row()] = data
        self._work_times.sort(key=lambda x: x["start"], reverse=False)
        self.endResetModel()

class OnCallDutyList(QAbstractListModel):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._events = list()

    def __iter__(self):
        return iter(self._events)

    def __getitem__(self, i):
        return self._events[i]

    def __len__(self):
        return len(self._events)

    def rowCount(self, parent=None):
        return len(self._events)

    def data(self, index, role):
        event = self._events[index.row()]
        if role == Qt.ItemDataRole.ToolTipRole or role == Qt.ItemDataRole.DisplayRole:
            return f"{event["start"].toString('dd.MM HH:mm')} - {event["end"].toString('dd.MM HH:mm')}\t{event["comments"]}"
        elif role == Qt.ItemDataRole.UserRole:
            return event

    def removeRow(self, index):
        self.beginRemoveRows(QModelIndex(), index, index)
        del self._events[index]
        self.endRemoveRows()

    def find(self, day_of_month):
        return [item for item in self._events if item["start"].date().day() == day_of_month] or None

    def clear(self):
        self.beginResetModel()
        self._events.clear()
        self.endResetModel()

    def addEvent(self, event):
        index = QModelIndex()
        self.beginInsertRows(index, self.rowCount(index), self.rowCount(index))
        self._events.append(event)
        self._events.sort(key=lambda x: x["start"], reverse=False)
        self.endInsertRows()

    def modifyEvent(self, index, data):
        self.beginResetModel()
        self._events[index.row()] = data
        self._events.sort(key=lambda x: x["start"], reverse=False)
        self.endResetModel()

    def setEvents(self, data):
        self.beginResetModel()
        self._events = [{'start': QDateTime.fromSecsSinceEpoch(item['start']), 'end': QDateTime.fromSecsSinceEpoch(item['end']), 'comments': item['comments']} for item in
                 data]
        self._events.sort(key=lambda x: x["start"], reverse=False)
        self.endResetModel()

    def getEvents(self):
        items = [{'start': item['start'].toSecsSinceEpoch(), 'end': item['end'].toSecsSinceEpoch(), 'comments': item['comments']} for item in
                 self._events]
        return items


class Workdays(QAbstractListModel):

    def __init__(self, workdays_spreadsheet, workdays_saved, month, year, parent=None):
        super().__init__(parent)
        self._workdays = list()
        for i, value in enumerate(workdays_spreadsheet):
            if workdays_saved is None:
                dict_item = {
                    "dayOfMonth": value["dayOfMonth"],
                    "dayOfWeek": value["dayOfWeek"],
                    "action": 0,  # working...vacation, sick, etc
                    "worktimes": WorktimeListModel()
                }
            else:
                value_from_saved = next((item for item in workdays_saved if item.get('dayOfMonth') == value["dayOfMonth"]), None)
                if value_from_saved is None or value["dayOfWeek"] != value_from_saved["dayOfWeek"]:
                    print("Something went wrong")
                    exit(1)
                dict_item = {
                    "dayOfMonth": value["dayOfMonth"],
                    "dayOfWeek": value["dayOfWeek"],
                    "action": value_from_saved['action'],  # working...vacation, sick, etc
                    "worktimes": WorktimeListModel(value_from_saved['worktimes'])
                }
            self._workdays.append(dict_item)

        self._month = month
        self._year = year

    def __iter__(self):
        return iter(self._workdays)

    def __getitem__(self, i):
        return self._workdays[i]

    def __len__(self):
        return len(self._workdays)

    def rowCount(self, parent=None):
        return len(self._workdays)

    def data(self, index, role):
        day = self._workdays[index.row()]
        if role == Qt.ItemDataRole.ToolTipRole or role == Qt.ItemDataRole.DisplayRole:
            return f"{day["dayOfMonth"]}\t{day["dayOfWeek"]}"
        elif role == Qt.ItemDataRole.UserRole:
            return day

    def setAction(self, index, action):
        self._workdays[index.row()]["action"] = action
        print(self._workdays[index.row()])

    def getWorktimeList(self, index):
        return self._workdays[index.row()]["worktimes"]

    def find(self, day_of_month):
        return next((item for item in self._workdays if item["dayOfMonth"] == day_of_month), None)

    def getData(self):
        items = [
            {
                'dayOfMonth': x['dayOfMonth'],
                'dayOfWeek': x['dayOfWeek'],
                'action': x['action'],
                'worktimes': x['worktimes'].getData()
            } for x in self._workdays]
        return items

class WorkTimeDialog(QDialog):
    def __init__(self, parent=None, initialData=None):
        super(WorkTimeDialog, self).__init__(parent)
        # load ui
        uic.loadUi("worktime.ui", self)
        self.startTimeEdit = self.findChild(QTimeEdit, "timeEditEventStart")
        self.endTimeEdit = self.findChild(QTimeEdit, "timeEditEventEnd")
        self.comboBoxWorkTypes = self.findChild(QComboBox, "comboBoxWorkTypes")
        self.comboBoxWorkTypes.addItems(WORKTYPES)

        if initialData is not None:
            self.startTimeEdit.setTime(initialData["start"])
            self.endTimeEdit.setTime(initialData["end"])
            self.comboBoxWorkTypes.setCurrentIndex(initialData["type"])

    def get_worktime(self):
        return dict({"start": self.startTimeEdit.time(), "end": self.endTimeEdit.time(), "type": self.comboBoxWorkTypes.currentIndex()})


class OnCallDutyDialog(QDialog):
    def __init__(self, parent=None, initialData=None):
        super(OnCallDutyDialog, self).__init__(parent)
        # load ui
        uic.loadUi("ocd.ui", self)
        self.startTimeEdit = self.findChild(QDateTimeEdit, "dateTimeEditEventStart")
        self.endTimeEdit = self.findChild(QDateTimeEdit, "dateTimeEditEventEnd")
        self.commentsEdit = self.findChild(QLineEdit, "lineEditComments")
        self.durationEdit = self.findChild(QLineEdit, "lineEditDuration")
        self.pushButtonCalcEventEnd = self.findChild(QPushButton, "pushButtonCalcEventEnd")
        self.pushButtonCalcEventEnd.clicked.connect(self.calculateEndTime)

        if initialData is None:
            self.startTimeEdit.setDateTime(QDateTime.currentDateTime())
        else:
            self.startTimeEdit.setDateTime(initialData["start"])
            self.endTimeEdit.setDateTime(initialData["end"])
            self.commentsEdit.setText(initialData["comments"])

    def get_ocd(self):
        ocd = dict()
        ocd['start'] = self.startTimeEdit.dateTime()
        ocd['end'] = self.endTimeEdit.dateTime()
        ocd['comments'] = self.commentsEdit.text()
        return ocd

    def calculateEndTime(self):
        startTime = self.startTimeEdit.dateTime()
        duration = int(self.durationEdit.text())
        seconds_to_add = duration * 60
        new_datetime = startTime.addSecs(seconds_to_add)
        self.endTimeEdit.setDateTime(new_datetime)


    def inputCheck(self):
        if self.endTimeEdit.dateTime() > self.startTimeEdit.dateTime():
            return True
        else:
            QMessageBox.information(None, "Warning!", "End time must be greater than start time.")
            return False

    def accept(self):
        validInput = self.inputCheck()
        if validInput:
            self.done(1)  # Only accept the dialog if all inputs are valid

class MainWindow(QMainWindow):

    def __init__(self,parent=None):
        super(MainWindow,self).__init__(parent)
        uic.loadUi("wt.ui", self)

        self.balance = {}

        self.firstNameEdit = self.findChild(QLineEdit, "lineEditFirstName")
        self.lastNameEdit = self.findChild(QLineEdit, "lineEditLastName")
        self.groupNameEdit = self.findChild(QLineEdit, "lineEditGroupName")
        self.targetMonthSpin = self.findChild(QSpinBox, "spinBoxTargetMonth")
        self.targetYearSpin = self.findChild(QSpinBox, "spinBoxTargetYear")
        self.workingPathEdit = self.findChild(QLineEdit, "lineEditWorkingPath")
        self.workingDaysList = self.findChild(QListView, "listViewWorkingDays")
        self.labelWorkdaysMonth = self.findChild(QLabel, "labelWorkdaysMonth")
        self.spinBoxBalanceHours = self.findChild(QSpinBox, "spinBoxBalanceHours")
        self.spinBoxBalanceMinutes = self.findChild(QSpinBox, "spinBoxBalanceMinutes")

        # Create models
        self.absence_items = QStringListModel(ACTIONS)
        weekday_items = QStringListModel(WEEKDAYS)
        self.customWorktimesModel = WorktimeListModel()  # useless?
        self.usualsModel = WeekdayUsualsList()
        self.ocdModel = OnCallDutyList()

        # Load resources, settings, etc
        self.loadSettings()
        self.loadBalanceConfiguration()
        self.loadOCD()

        self.targetYearSpin.valueChanged.connect(self.targetChanged)
        self.targetMonthSpin.valueChanged.connect(self.targetChanged)

        self.spinBoxBalanceHours.valueChanged.connect(self.balanceChanged)
        self.spinBoxBalanceMinutes.valueChanged.connect(self.balanceChanged)

        self.listViewActions = self.findChild(QListView, "listViewActions")
        self.listViewActions.setModel(self.absence_items)
        self.listViewActions.selectionModel().selectionChanged.connect(self.actionChanged)

        self.listViewUsualWeekdays = self.findChild(QListView, "listViewUsualWeekdays")
        self.listViewUsualWeekdays.setModel(weekday_items)
        self.listViewUsualWeekdays.selectionModel().selectionChanged.connect(self.usualsChanged)

        self.listViewWorktimes = self.findChild(QListView, "listViewWorktimes")
        self.listViewWorktimes.setModel(self.customWorktimesModel)
        self.listViewWorktimes.doubleClicked.connect(self.editWorktime)

        self.listViewWorktimeUsual = self.findChild(QListView, "listViewWorktimeUsual")
        self.listViewWorktimeUsual.setModel(self.usualsModel)
        self.listViewWorktimeUsual.doubleClicked.connect(self.editUsual)
        self.loadUsuals()

        self.listViewOCD = self.findChild(QListView, "listViewOCD")
        self.ocdModel.rowsInserted.connect(lambda: self.saveOCD())
        self.ocdModel.rowsRemoved.connect(lambda: self.saveOCD())
        self.listViewOCD.setModel(self.ocdModel)
        self.listViewOCD.doubleClicked.connect(self.editOCD)

        self.pushButtonUpdateWorkdays = self.findChild(QPushButton, "pushButtonUpdateWorkdays")
        self.pushButtonUpdateWorkdays.clicked.connect(lambda: self.updateWorkdays())

        self.pushButtonAddOCD = self.findChild(QPushButton, "pushButtonAddOCD")
        self.pushButtonAddOCD.clicked.connect(lambda: self.addOCD())

        self.pushButtonRemoveOCD = self.findChild(QPushButton, "pushButtonRemoveOCD")
        self.pushButtonRemoveOCD.clicked.connect(lambda: self.removeOCD())

        self.pushButtonSelectWorkingDir = self.findChild(QPushButton, "pushButtonSelectWorkingDir")
        self.pushButtonSelectWorkingDir.clicked.connect(lambda: self.selectWorkingDir())

        self.pushButtonAddWorktime = self.findChild(QPushButton, "pushButtonAddWorktime")
        self.pushButtonAddWorktime.clicked.connect(lambda: self.addWorktime())

        self.pushButtonRemoveWorktime = self.findChild(QPushButton, "pushButtonRemoveWorktime")
        self.pushButtonRemoveWorktime.clicked.connect(lambda: self.removeWorktime())

        self.pushButtonAddWorktimeUsual = self.findChild(QPushButton, "pushButtonAddWorktimeUsual")
        self.pushButtonAddWorktimeUsual.clicked.connect(lambda: self.addWorktimeUsual())
        self.pushButtonRemoveWorktimeUsual = self.findChild(QPushButton, "pushButtonRemoveWorktimeUsual")
        self.pushButtonRemoveWorktimeUsual.clicked.connect(lambda: self.removeWorktimeUsual())

        self.pushButtonCreateSpreadsheet = self.findChild(QPushButton, "pushButtonCreateSpreadsheet")
        self.pushButtonCreateSpreadsheet.clicked.connect(lambda: self.createSpreadsheet())

        self.statusBar().showMessage('Application is initialized')


    def editWorktime(self, item=None):
        data = self.customWorktimesModel.data(item, role=Qt.ItemDataRole.UserRole)
        dialog = WorkTimeDialog(self, data)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print(f"Editing worktime: {dialog.get_worktime()}")
            self.customWorktimesModel.modifyItem(item, dialog.get_worktime())
        else:
            print("Editing worktime cancelled")

    def editOCD(self, item=None):
        data = self.ocdModel.data(item, role=Qt.ItemDataRole.UserRole)
        dialog = OnCallDutyDialog(self, data)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print(f"Editing OCD: {dialog.get_ocd()}")
            self.ocdModel.modifyEvent(item, dialog.get_ocd())
        else:
            print("Editing OCD cancelled")

    def editUsual(self, item=None):
        data = self.usualsModel.data(item, role=Qt.ItemDataRole.UserRole)
        dialog = WorkTimeDialog(self, data)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print(f"Editing usuals: {dialog.get_worktime()}")
            self.usualsModel.modify_work_time(item, dialog.get_worktime())
        else:
            print("Editing usuals cancelled")

    def saveSetting(self):
        print("Save settings")
        self.settings.setValue("firstName", self.firstNameEdit.text())
        self.settings.setValue("lastName", self.lastNameEdit.text())
        self.settings.setValue("groupName", self.groupNameEdit.text())
        self.settings.setValue("targetMonth", self.targetMonthSpin.value())
        self.settings.setValue("targetYear", self.targetYearSpin.value())
        self.settings.setValue("workingPath", self.workingPathEdit.text())

    def loadSettings(self):
        self.settings = QSettings("./Settings.ini", QSettings.Format.IniFormat)
        try:
            print("Load settings...")
            self.firstNameEdit.setText(self.settings.value("firstName", "John", type=str))
            self.lastNameEdit.setText(self.settings.value("lastName", "Doe", type=str))
            self.groupNameEdit.setText(self.settings.value("groupName", "Black Magic", type=str))
            self.targetMonthSpin.setValue(self.settings.value("targetMonth", 9, type=int))
            self.targetYearSpin.setValue(self.settings.value("targetYear", 2024, type=int))
            self.workingPathEdit.setText(self.settings.value("workingPath", "", type=str))
        except:
            pass

    def loadBalanceConfiguration(self):
        if not os.path.isfile('balance.json'):
            with open('balance.json', 'w') as f:
                json.dump(dict(), f)
        else:
            with open('balance.json', 'r') as f:
                self.balance = json.load(f)
            key = f"{self.targetMonthSpin.value()}.{self.targetYearSpin.value()}"
            if key in self.balance:
                with QSignalBlocker(self.spinBoxBalanceHours):
                    self.spinBoxBalanceHours.setValue(self.balance[key]['h'])
                with QSignalBlocker(self.spinBoxBalanceMinutes):
                    self.spinBoxBalanceMinutes.setValue(self.balance[key]['m'])
            else:
                with QSignalBlocker(self.spinBoxBalanceHours):
                    self.spinBoxBalanceHours.setValue(0)
                with QSignalBlocker(self.spinBoxBalanceMinutes):
                    self.spinBoxBalanceMinutes.setValue(0)

    def closeEvent(self, event):
        self.saveSetting()
        self.saveUsuals()
        self.saveWorktimes()
        self.saveBalance()
        print("Exit")

    def balanceChanged(self):
        h = self.spinBoxBalanceHours.value()
        m = self.spinBoxBalanceMinutes.value()
        key = f"{self.targetMonthSpin.value()}.{self.targetYearSpin.value()}"
        self.balance[key] = dict({"h": h, "m": m})
        print(f"balanceChanged {h}:{m}")


    def targetChanged(self, item):
        self.loadOCD()

        print(f"targetChanged {item}")
        print(f"\tbalance: {self.balance}")
        key = f"{self.targetMonthSpin.value()}.{self.targetYearSpin.value()}"
        print(f"\tkey: {key}")
        if key in self.balance:
            print(f"\ttargetChanged with key {self.balance[key]['h']}:{self.balance[key]['m']}")
            with QSignalBlocker(self.spinBoxBalanceHours):
                self.spinBoxBalanceHours.setValue(self.balance[key]['h'])
            self.spinBoxBalanceMinutes.setValue(self.balance[key]['m'])
        else:
            print(f"\ttargetChanged with key zero")
            with QSignalBlocker(self.spinBoxBalanceHours):
                self.spinBoxBalanceHours.setValue(0)
            self.spinBoxBalanceMinutes.setValue(0)

    def saveWorktimes(self):
        try:
            with open(f'worktimes-{self.current_target_month}-{self.current_target_year}.json', 'w') as f:
                json.dump(self.workDaysModel.getData(), f)
        except:
            print("Worktimes not saved...")
            pass

    def saveOCD(self):
        with open(f'ocd-{self.targetMonthSpin.value()}-{self.targetYearSpin.value()}.json', 'w') as f:
            json.dump(self.ocdModel.getEvents(), f)
        print("Saving OCD")

    def saveBalance(self):
        with open('balance.json', 'w') as f:
            json.dump(self.balance, f)
        print("Saving balance")

    def saveUsuals(self):
        with open('usuals.json', 'w') as f:
            json.dump(self.usualsModel.getUsuals(), f)

    def loadOCD(self):
        fn = f'ocd-{self.targetMonthSpin.value()}-{self.targetYearSpin.value()}.json'
        print(fn)
        if os.path.isfile(fn):
            try:
                with open(fn, 'r') as f:
                    self.ocdModel.setEvents(json.load(f))
                    print("Loaded OCD")
            except Exception:
                print("Error loading OCD")
                pass
        else:
            self.ocdModel.clear()

    def loadUsuals(self):
        try:
            with open('usuals.json', 'r') as f:
                self.usualsModel.setUsuals(json.load(f))
        except Exception:
            pass

    def loadWorktimes(self):
        fn = f'worktimes-{self.targetMonthSpin.value()}-{self.targetYearSpin.value()}.json'
        if os.path.isfile(fn):
            try:
                with open(fn, 'r') as f:
                    return json.load(f)
            except Exception:
                return None
        else:
            return None

    def addWorktimeUsual(self):
        dialog = WorkTimeDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print("Adding usuals")
            self.usualsModel.add_work_time(dialog.get_worktime())
        else:
            print("Adding usuals cancelled")

    def removeWorktimeUsual(self):
        reply = QMessageBox.question(self, "Message", "Really remove selected usual worktime?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Yes:
            row = self.listViewWorktimeUsual.selectionModel().currentIndex().row()
            print(f"Removing usuals {row}")
            self.listViewWorktimeUsual.model().removeRow(row)

    def addOCD(self):
        dialog = OnCallDutyDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print(f"Adding OCD: {dialog.get_ocd()}")
            self.ocdModel.addEvent(dialog.get_ocd())
        else:
            print("Adding OCD cancelled")

    def removeOCD(self):
        reply = QMessageBox.question(self, "Message", "Really remove selected ocd?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Yes:
            row = self.listViewOCD.selectionModel().currentIndex().row()
            print(f"Removing OCD {row}")
            self.listViewOCD.model().removeRow(row)

    def selectWorkingDir(self):
        working_path = QFileDialog.getExistingDirectory(self, 'Select Folder')
        self.workingPathEdit.setText(working_path)
        print(working_path)

    def workingDayChanged(self, selected_item, deselected_item):
        if selected_item.indexes():
            item = self.workDaysModel.data(selected_item.indexes()[0], Qt.ItemDataRole.UserRole)
            print(f"workingDayChanged: {selected_item.indexes()[0].row()} - {item}")
            # action list view
            self.listViewActions.setEnabled(True)
            self.listViewActions.selectionModel().clear()
            self.listViewActions.selectionModel().setCurrentIndex(self.absence_items.index(item["action"]), QItemSelectionModel.SelectionFlag.Select)
            # worktimes list view
            self.listViewWorktimes.setModel(item["worktimes"])
            self.customWorktimesModel = item["worktimes"]


    def actionChanged(self, selected_item, deselected_item):
        if selected_item.indexes():
            action_row = selected_item.indexes()[0].row()
            self.workDaysModel.setAction(self.workingDaysList.selectionModel().currentIndex(), action_row)
            # enable disable worktime recording
            if action_row == 0:
                self.listViewWorktimes.setEnabled(False)
                self.pushButtonAddWorktime.setEnabled(False)
                self.pushButtonRemoveWorktime.setEnabled(False)
            elif action_row == 1:
                self.listViewWorktimes.setEnabled(True)
                self.pushButtonAddWorktime.setEnabled(True)
                self.pushButtonRemoveWorktime.setEnabled(True)
            elif action_row == 2:
                self.listViewWorktimes.setEnabled(False)
                self.pushButtonAddWorktime.setEnabled(False)
                self.pushButtonRemoveWorktime.setEnabled(False)
            elif action_row == 3:
                self.listViewWorktimes.setEnabled(True)
                self.pushButtonAddWorktime.setEnabled(True)
                self.pushButtonRemoveWorktime.setEnabled(True)
            elif action_row == 4:
                self.listViewWorktimes.setEnabled(False)
                self.pushButtonAddWorktime.setEnabled(False)
                self.pushButtonRemoveWorktime.setEnabled(False)
            elif action_row == 5:
                self.listViewWorktimes.setEnabled(True)
                self.pushButtonAddWorktime.setEnabled(True)
                self.pushButtonRemoveWorktime.setEnabled(True)
            elif action_row == 6:
                self.listViewWorktimes.setEnabled(True)
                self.pushButtonAddWorktime.setEnabled(True)
                self.pushButtonRemoveWorktime.setEnabled(True)

    def addWorktime(self):
        dialog = WorkTimeDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            print(f"Adding work time: {dialog.get_worktime()}")
            self.listViewWorktimes.model().addItem(dialog.get_worktime())
        else:
            print("Adding work time cancelled")

    def removeWorktime(self):
        reply = QMessageBox.question(self, "Message", "Really remove selected worktime?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Yes:
            row = self.listViewWorktimes.selectionModel().currentIndex().row()
            print(f"Removing worktime {row}")
            self.listViewWorktimes.model().removeRow(row)


    def usualsChanged(self, selected_item, deselected_item):
        if selected_item.indexes():
            index = selected_item.indexes()[0].row()
            self.listViewWorktimeUsual.setEnabled(True)
            self.pushButtonAddWorktimeUsual.setEnabled(True)
            self.pushButtonRemoveWorktimeUsual.setEnabled(True)
            self.usualsModel.set_weekday(index)
            print(WEEKDAYS[index])

    def createSpreadsheet(self):
        search_pattern = os.path.join(self.workingPathEdit.text(), 'LastName_FirstName_*.xlsx')
        matching_files = glob.glob(search_pattern)
        if len(matching_files) != 1:
            print("No templates found")  # write in status bar
            return
        template_file = matching_files[0]
        try:
            workbook = xw.Book(template_file)
            worksheet_profile = workbook.sheets['My Profile']
            worksheet_plan = workbook.sheets['Monthly Plan and Absences']
            worksheet_time = workbook.sheets['Enter Working Time']
            worksheet_record = workbook.sheets['Work Time Record']

            # Change the target month
            worksheet_plan.range('C5').value = self.targetMonthSpin.value()

            # Profile
            worksheet_profile.range('C3').value = f'{self.firstNameEdit.text()} {self.lastNameEdit.text()}'
            worksheet_profile.range('C4').value = self.groupNameEdit.text()

            # iterate through all days in target month
            month_tuple = monthrange(self.targetYearSpin.value(), self.targetMonthSpin.value())
            days_in_month = month_tuple[1]
            worktime_row = WORKTIME_STARTING_ROW
            for day_of_month in range(1, days_in_month + 1):
                # try to find the day in workdays
                workday = self.workDaysModel.find(day_of_month)
                if workday is not None:
                    action = workday['action']
                    if action >= 2:
                        # neither work nor ocd is possible here
                        worksheet_plan.range(f'{PLAN_ABSENCE_COL}{PLAN_STARTING_ROW + day_of_month - 1}').value = ACTIONS[action]
                    else:
                        if action == 0:  # usuals
                            day_of_week_index = WEEKDAYS.index(workday["dayOfWeek"])
                            usuals = self.usualsModel.find(day_of_week_index)
                            for u in usuals:
                                worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row}').value = WORKTYPES[u['type']]
                                worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row}').value = day_of_month
                                worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row}').value = day_of_month
                                worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row}').value = u['start'].toString("HH:mm")
                                worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row}').value = u['end'].toString("HH:mm")
                                worktime_row += 1
                        else:  # custom times
                            custom_times = workday["worktimes"].getWorkTimes()
                            for c in custom_times:
                                worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row}').value = WORKTYPES[c['type']]
                                worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row}').value = day_of_month
                                worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row}').value = day_of_month
                                worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row}').value = c['start'].toString("HH:mm")
                                worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row}').value = c['end'].toString("HH:mm")
                                worktime_row += 1

                # check if there is OCD on that day
                ocd = self.ocdModel.find(day_of_month)
                if ocd is not None:
                    for o in ocd:
                        worksheet_time.range(f'{WORKTIME_TYPE_COL}{worktime_row}').value = 'OCD'
                        worksheet_time.range(f'{WORKTIME_START_DAY_COL}{worktime_row}').value = o["start"].date().day()
                        worksheet_time.range(f'{WORKTIME_END_DAY_COL}{worktime_row}').value = o["end"].date().day()
                        worksheet_time.range(f'{WORKTIME_START_TIME_COL}{worktime_row}').value = o["start"].time().toString("HH:mm")
                        worksheet_time.range(f'{WORKTIME_END_TIME_COL}{worktime_row}').value = o["end"].time().toString("HH:mm")
                        worksheet_time.range(f'{WORKTIME_COMMENTS_COL}{worktime_row}').value = o["comments"]
                        worktime_row += 1

            # Save the workbook with a new name
            fn = f'{self.lastNameEdit.text()}_{self.firstNameEdit.text()}_WorkTimeRecord_{self.targetYearSpin.value()}-{self.targetMonthSpin.value():02}.xlsx'
            fn_with_path = os.path.join(self.workingPathEdit.text(), fn)

            workbook.save(fn_with_path)

            # find balance
            for row in range(1, 50):
                cell_value = worksheet_record.range(f'T{row}').value
                if isinstance(cell_value, str) and "balance" in cell_value:
                    balance_combined = worksheet_record.range(f'W{row}').value
                    print(balance_combined)
                    balance_h = int(balance_combined.split(":")[0])
                    balance_m = int(balance_combined.split(":")[1])
                    d = QDate(self.current_target_year, self.current_target_month, 1)
                    d = d.addMonths(1)
                    key = f"{d.month()}.{d.year()}"
                    self.balance[key] = dict({"h": balance_h, "m": balance_m})
                    break

            app = workbook.app
            workbook.close()
            app.kill()

        except Exception as e:
            QMessageBox.critical(None, "Error reading template", str(e))
            raise e

    def updateWorkdays(self):
        search_pattern = os.path.join(self.workingPathEdit.text(), 'LastName_FirstName_*.xlsx')
        matching_files = glob.glob(search_pattern)
        if len(matching_files) != 1:
            print("No templates found")  # write in status bar
            return
        template_file = matching_files[0]
        try:
            workbook = xw.Book(template_file)
            worksheet_plan = workbook.sheets['Monthly Plan and Absences']
            # Change the target month
            worksheet_plan.range('C5').value = self.targetMonthSpin.value()
            working_days = []
            for row in range(PLAN_STARTING_ROW, PLAN_STARTING_ROW + 31):
                day_type = worksheet_plan.range(f'{PLAN_DAYTYPE_COL}{row}').value
                week_day = worksheet_plan.range(f'{PLAN_WEEKDAY_COL}{row}').value
                if day_type == 'Working day':
                    working_days.append({"dayOfMonth": int(worksheet_plan.range(f'{PLAN_DAYOFMONTH_COL}{row}').value),
                                         "dayOfWeek": week_day})

            self.workDaysModel = Workdays(working_days, self.loadWorktimes(), self.targetMonthSpin.value(), self.targetYearSpin.value())

            self.workingDaysList.setModel(self.workDaysModel)
            self.workingDaysList.selectionModel().selectionChanged.connect(self.workingDayChanged)

            self.labelWorkdaysMonth.setText(f"{self.targetMonthSpin.value()}.{self.targetYearSpin.value()}")
            self.current_target_month = self.targetMonthSpin.value()
            self.current_target_year = self.targetYearSpin.value()

            app = workbook.app
            workbook.close()
            app.kill()
        except Exception as e:
            QMessageBox.critical(None, "Error reading template", str(e))
            raise e



def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    app.exec()


if __name__ == '__main__':
    main()

