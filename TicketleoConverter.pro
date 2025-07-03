QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    main.cpp

HEADERS +=

FORMS +=

RESOURCES += \
    resources.qrc

COMMIT = '\\"$$system(git describe --tags --dirty)\\"'
DEFINES += COMMIT_VERSION=\"$${COMMIT}\"

# QXlsx code for Application Qt project
QXLSX_PARENTPATH=./QXlsx/QXlsx
QXLSX_HEADERPATH=./QXlsx/QXlsx/header/
QXLSX_SOURCEPATH=./QXlsx/QXlsx/source/
include(./QXlsx/QXlsx/QXlsx.pri)

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target
