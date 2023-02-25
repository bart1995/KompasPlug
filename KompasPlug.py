import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D

def kompas_plug(KompasType):
	ModuleConstants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
	Module5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
	Module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
	kompas_object = Module5.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(Module5.KompasObject.CLSID, 
												      pythoncom.IID_IDispatch))
	application = Module7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(Module7.IApplication.CLSID, 
												    pythoncom.IID_IDispatch))
	Documents = application.Documents
	KompasType2D = {	'Drawing':ModuleConstants.ksDocumentDrawing, 
				'Fragment':ModuleConstants.ksDocumentFragment, 
				'Textual':ModuleConstants.ksDocumentTextual, 
				'Specification':ModuleConstants.ksDocumentSpecification}
	KompasType3D = {'Model3D':ModuleConstants.ksDocumentPart, 'Assembly':ModuleConstants.ksDocumentAssembly}
	if KompasType in KompasType2D:
		kompas_document = Documents.AddWithDefaultSettings(KompasType2D[KompasType], True)
		kompas_document_2d = Module7.IKompasDocument2D(kompas_document)
		iDocument = kompas_object.ActiveDocument2D()
	elif KompasType in KompasType3D:
		kompas_document = Documents.AddWithDefaultSettings(KompasType3D[KompasType], True)
		kompas_document_3d = Module7.IKompasDocument3D(kompas_document)
		iDocument = kompas_object.ActiveDocument3D()
	elif KompasType == 'ActiveDocument2D':
		kompas_document = application.ActiveDocument
		kompas_document_2d = Module7.IKompasDocument2D(kompas_document)
		iDocument = kompas_object.ActiveDocument2D()
	elif KompasType == 'ActiveDocument3D':
		kompas_document = application.ActiveDocument
		kompas_document_3d = Module7.IKompasDocument3D(kompas_document)
		iDocument = kompas_object.ActiveDocument3D()
	else:
		print('Программа не имеет такого модуля')
	return Module5, Module7, iDocument
