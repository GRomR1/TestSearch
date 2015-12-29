#-------------------------------------------------
#
# Project created by QtCreator 2015-12-25T16:00:01
#
#-------------------------------------------------

QT       += core gui sql axcontainer concurrent
CONFIG += c++11
QMAKE_CXXFLAGS += -std=c++11

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = TestSearch
TEMPLATE = app


SOURCES += address.cpp \
    database.cpp \
    xlsparser.cpp

HEADERS  +=  defines.h \
    address.h \
    database.h \
    xlsparser.h

SOURCES += main.cpp\
        widget.cpp \
    excel.cpp

HEADERS  += widget.h \
    excel.h

FORMS    += widget.ui
