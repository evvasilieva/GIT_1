# -*- coding: utf-8 -*-
#|макрос

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH

kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
MH.iKompasObject  = kompas_object

kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
MH.iApplication  = application


Documents = application.Documents
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

Excel=Dispatch('Excel.Application')
book1=Excel.Workbooks.Open("F:\parametr.xlsx", ReadOnly=True)
sheet1 = book1.Sheets('Заказ')
sheet2 = book1.Sheets('Отступы')

i_zakaz = 1
delta= 0

while True:
    i_zakaz = i_zakaz + 1
    if i_zakaz > 100:
        break
    profil = sheet1.Cells(i_zakaz,1).value
    if profil is None:
        break
    q = int(sheet1.Cells(i_zakaz,4).value)
    w = sheet1.Cells(i_zakaz,3).value
    h = sheet1.Cells(i_zakaz,2).value

    for num in range(0,q):

        k = 1
        while True:
            k = k + 1
            if k > 1000:
                break
            name = sheet2.Cells(k,1).value
            otst = sheet2.Cells(k,2).value
            if name is None:
                break
            if name == profil:
                print(profil,h,w,q,otst)


                # dfgdf g
                iRectangleParam = kompas6_api5_module.ksRectangleParam(kompas_object.GetParamStruct(kompas6_constants.ko_RectangleParam))
                iRectangleParam.Init()
                iRectangleParam.x = delta + otst
                iRectangleParam.y = otst
                iRectangleParam.ang = 0
                iRectangleParam.height = h - 2 * otst
                iRectangleParam.width = w - 2 * otst
                iRectangleParam.style = 1
                obj = iDocument2D.ksRectangle(iRectangleParam)

        delta = delta + w + 50

