import sys
import subprocess
import logging
import time
import win32gui
import win32process
import argparse

from PyQt4 import QtGui,  QtCore
from PyQt4.QtCore import Qt

try:
    import config
except ImportError:
    print("File config.py missing. Maybe you have to rename config.template?")
    sys.exit(1)
    
PROGNAME = 'launcher'
VERSION = 'V1.0.0'

def get_hwnd_from_pid(pid):
    handles = []
    def _callback(hwnd, pid):
        (threadId, processId) = win32process.GetWindowThreadProcessId(hwnd)
        if ((processId == pid) or (threadId == pid)) and win32gui.GetParent(hwnd) == 0:
            handles.append(hwnd)
    win32gui.EnumWindows(_callback, pid)
    if len(handles) > 0:
        return handles[0]
    else:
        return 0
# ------------------------------------------------------------------------------
class DDPushButton(QtGui.QPushButton):
    def __init__(self, caption):
        super(DDPushButton, self).__init__(caption)
        self.data = ""
        self.config = {}
        self.setAcceptDrops(True)

    def setDropable(self, bool):
        self.setAcceptDrops(bool)

    def setConfig(self, config):
        self.config = config
        if "icon" in config:
            self.setIcon(QtGui.QIcon(config["icon"]))

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("text/uri-list"):
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.data = event.mimeData().urls()
        if self.config["history"] == True:
            cbHistory = self.config["__cbHistory"]
            for url in self.data:
                cbHistory.insertItem(0, url.toLocalFile())
                if cbHistory.count() > 10:
                    cbHistory.removeItem(cbHistory.count() - 1)
            cbHistory.setCurrentIndex(0)
        event.acceptProposedAction()
        self.clicked.emit(True)
        self.data = ""


class DDLabel(QtGui.QLabel):
    def __init__(self, caption):
        super(DDLabel, self).__init__(caption)
        self.pushButton = None
        self.config = {}
        self.setAcceptDrops(True)
        self.setTextInteractionFlags(QtCore.Qt.LinksAccessibleByMouse | QtCore.Qt.LinksAccessibleByKeyboard)
        self.setToolTip("No help available")
                    

    def setPushButton(self, pushButton):
        self.pushButton = pushButton

    def setConfig(self, config):
        self.config = config

    def dragEnterEvent(self, event):
        self.pushButton.dragEnterEvent(event)

    def dropEvent(self, event):
        self.pushButton.dropEvent(event)


class DDComboBox(QtGui.QComboBox):
    def __init__(self):
        super(DDComboBox, self).__init__()
        self.pushButton = None
        self.setAcceptDrops(True)

    def setPushButton(self, pushButton):
        self.pushButton = pushButton

    def dragEnterEvent(self, event):
        self.pushButton.dragEnterEvent(event)

    def dropEvent(self, event):
        self.pushButton.dropEvent(event)


class MainWin(QtGui.QDialog):
    """Application main window
    """
    def __init__(self, app):
        """Setup the window
        """
        super(MainWin, self).__init__()
        self.setWindowTitle(PROGNAME)
        self.setMinimumSize(400, 0)
        self.setWindowFlags(Qt.Dialog)
        self.stdwinapp = "-n" in app.arguments()
        self.minimizeAction = QtGui.QAction("Mi&nimize", self, triggered=self.hide)
        self.maximizeAction = QtGui.QAction("Ma&ximize", self, triggered=self.showMaximized)
        self.restoreAction = QtGui.QAction("&Restore", self, triggered=self.showNormal)
        quitAction = QtGui.QAction("&Quit", self, triggered=app.quit)
        aboutAction = QtGui.QAction("&About {0}".format(PROGNAME), self, triggered=self.aboutMe)
        aboutQtAction = QtGui.QAction("&About Qt", self, triggered=self.aboutQt)
        clearSettingsAction = QtGui.QAction("&Clear Settings", self, triggered=self.clearSettings)

        self.icon = QtGui.QIcon("icon_16x16.png")
        self.setWindowIcon(self.icon)

        trayIconMenu = QtGui.QMenu(self)
        trayIconMenu.addAction(self.minimizeAction)
        trayIconMenu.addAction(self.maximizeAction)
        trayIconMenu.addAction(self.restoreAction)
        trayIconMenu.addSeparator()
        trayIconMenu.addAction(clearSettingsAction)
        trayIconMenu.addAction(aboutAction)
        trayIconMenu.addAction(aboutQtAction)
        trayIconMenu.addSeparator()
        trayIconMenu.addAction(quitAction)

        self.trayIcon = QtGui.QSystemTrayIcon(self)
        self.trayIcon.setIcon(self.icon)
        self.trayIcon.setToolTip("Launcher")
        self.trayIcon.setContextMenu(trayIconMenu)
        self.trayIcon.activated.connect(self.iconActivated)
        self.trayIcon.show()

        tabWidget = QtGui.QTabWidget()
        for cfg in config.CONFIGLIST:
            appWidget = QtGui.QWidget()
            layout = QtGui.QGridLayout()
            for button, row in zip(cfg["buttons"], range(len(cfg["buttons"]))):
                pushButton = DDPushButton("Run")
                pushButton.setDropable(button.get("draggable", False))
                pushButton.setConfig(button)
                pushButton.clicked.connect(self.buttonClicked)
                layout.addWidget(pushButton, row, 0)
                if button.get("help") or button.get("manual"):
                    text = "<a href='{0}'>{0}</a>".format(button["name"])
                else:
                    text = button["name"]
                label = DDLabel(text)
                label.setConfig(button)
                if button.get("help"):
                    if isinstance(button.get("help"), type("")):
                        label.setToolTip(button["help"])
                    else:            
                        settings = QtCore.QSettings()
                        settings.beginGroup("HELP_" + cfg["tabname"])
                        if not button["name"] in settings.childKeys():
                            # if help string is not in settings file add it 
                            logging.debug("get help and write to settings")
                            settings.setValue(button["name"], button.get("help")[0](button.get("help")[1]))
                        label.setToolTip(settings.value(button["name"]).toString())
                        settings.endGroup()
                label.linkActivated.connect(self.showManual)
                label.setPushButton(pushButton)
                layout.addWidget(label, row, 1)
                if button.get("history", False):
                    cbHistory = DDComboBox()
                    cbHistory.setPushButton(pushButton)
                    button["__cbHistory"] = cbHistory
                    layout.addWidget(cbHistory, row, 2)
                    button["__cbHistory"] = cbHistory
            layout.setColumnStretch(2,2)
            appWidget.setLayout(layout)
            tabWidget.addTab(appWidget, cfg["tabname"])
        self.readSettings()
        layout = QtGui.QVBoxLayout()
        layout.addWidget(tabWidget)
        self.setLayout(layout)
        try:
            (x, y, width,  height) = config.POSITION
            self.setGeometry(x, y, width,  height)
        except AttributeError:
            pass


    def showManual(self, link):
        logging.debug("showManual")
        cmd = self.sender().config.get("manual")
        if isinstance(cmd, type([])) and len(cmd) == 2:
            cmd[0](cmd[1])


    def buttonClicked(self):
        sender = self.sender()
        data = sender.data
        if data == "":
            if sender.config.get("history", False) == True and sender.config["__cbHistory"].count() > 0:
                data = sender.config["__cbHistory"].itemText(0)
        else:
            data = data[0].toLocalFile().replace("/", "\\")
        for command in sender.config["commands"]:
            position = None
            size = None
            if isinstance(command, dict):
                position = command.get("position", None)
                size = command.get("size", None)
                command = command["command"]
            runargs = '{0} "{1}"'.format(command, data)
            logging.debug(runargs)
            pid = subprocess.Popen(runargs).pid
            time.sleep(0.5)
            if (position or size):
                hwnd = get_hwnd_from_pid(pid)
                logging.debug("{0}, {1}, {2}".format(position, size, hwnd))
                if hwnd == 0:
                    logging.debug("hwnd=0")
                    continue
                (left, top, right, bottom) = win32gui.GetWindowRect(hwnd)
                if position != None:
                    (left, top) = position
                if size != None:
                    (right, bottom) = (left + size[0], top + size[1])
                win32gui.SetWindowPos(hwnd, 0, left, top, right-left, bottom-top, 0)

    def readSettings(self):
        settings = QtCore.QSettings()
        for cfg in config.CONFIGLIST:
            settings.beginGroup(cfg["tabname"])
            for button in cfg["buttons"]:
                if button.get("history", False) == False:
                    continue;
                cbHistory = button["__cbHistory"]
                size = settings.beginReadArray("files");
                for i in range(size):
                    settings.setArrayIndex(i);
                    item = settings.value(button["name"], None)
                    if item != None:
                        cbHistory.addItem(item)
                settings.endArray()
            settings.endGroup()

    def writeSettings(self):
        settings = QtCore.QSettings()
        for cfg in config.CONFIGLIST:
            settings.beginGroup(cfg["tabname"])
            for button in cfg["buttons"]:
                if button.get("history", False) == False:
                    continue;
                cbHistory = button["__cbHistory"]
                settings.beginWriteArray("files");
                for i in range(cbHistory.count()):
                    settings.setArrayIndex(i);
                    settings.setValue(button["name"], cbHistory.itemText(i))
                settings.endArray()
            settings.endGroup()

    def clearSettings(self):
        logging.debug("clearSettings")
        settings = QtCore.QSettings()
        for key in settings.allKeys():
            settings.remove(key)
        for cfg in config.CONFIGLIST:
            for button in cfg["buttons"]:
                try:
                    if button["history"] != False:
                        button["__cbHistory"].clear()
                except KeyError:
                    pass

    def aboutMe(self):
        QtGui.QMessageBox.about(self, "About traylauncher", 
            "<b>traylauncher</b><br/><br/>" +
            "Program to lauch configurable programs with optional arguments via " + 
            "drag and drop or from history.<br/>" +
            "Configure via the file config.py<br/><br/>" + 
            "(C) 2013 Achim Koehler")

    def aboutQt(self):
        QtGui.QMessageBox.aboutQt(self)

    def setVisible(self, visible):
       self.minimizeAction.setEnabled(visible);
       self.maximizeAction.setEnabled(not self.isMaximized())
       self.restoreAction.setEnabled(self.isMaximized() or not visible);
       super(MainWin, self).setVisible(visible)

    def closeEvent(self, event):
        self.writeSettings()
        if (self.trayIcon.isVisible()):
            self.hide();
        if self.stdwinapp:
            event.accept()
        else:
            event.ignore()

    def iconActivated(self, reason):
        if (reason == QtGui.QSystemTrayIcon.Trigger) or (reason == QtGui.QSystemTrayIcon.DoubleClick):
            self.setVisible(not self.isVisible())
            if self.isVisible():
                self.showNormal()
        elif (reason == QtGui.QSystemTrayIcon.MiddleClick):
            pass


def start(args):
    """Start the application
    """
    app = QtGui.QApplication(sys.argv)
    app.setOrganizationName("")
    app.setOrganizationDomain("macht-publik.de")
    app.setApplicationName("launcher")
    QtCore.QSettings.setDefaultFormat(QtCore.QSettings.IniFormat)
    qt_translator = QtCore.QTranslator()
    qt_translator.load("qt_" + QtCore.QLocale.system().name(),
        QtCore.QLibraryInfo.location(QtCore.QLibraryInfo.TranslationsPath))
    app.installTranslator(qt_translator)
    if not QtGui.QSystemTrayIcon.isSystemTrayAvailable():
        QtGui.QMessageBox.critical(0, PROGNAME,
            "I couldn't detect any system tray on this system.")
        sys.exit(1)
    QtGui.QApplication.setQuitOnLastWindowClosed(args.notray)
    logging.debug("stdwinapp={0}".format(args.notray))
    mainwin = MainWin(app)
    mainwin.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    parser = argparse.ArgumentParser(description='Yet another program launcher')
    parser.add_argument('-n', dest='notray', action='store_true',
                   help='do not minimize to tray')
    args = parser.parse_args()
    start(args)