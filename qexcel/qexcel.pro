CONFIG += qaxcontainer

QT       += core widgets

QT       -= gui

TARGET = QExcel
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += main.cpp \
	qexcel.cpp

HEADERS += \
	qexcel.h

OTHER_FILES += \
	../excelengine.h \
	../excelengine.cpp
