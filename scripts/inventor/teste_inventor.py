import win32com.client
import pythoncom

pythoncom.CoInitialize()

inv = win32com.client.Dispatch("Inventor.Application")

print("OK - abriu Inventor")