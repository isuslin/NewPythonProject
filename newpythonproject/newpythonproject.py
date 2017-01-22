# -*- coding: utf-8 -*-

# To change this license header, choose License Headers in Project Properties.
# To change this template file, choose Tools | Templates
# and open the template in the editor.

import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.visible = 1
excel.workbooks.add(2)
excel.sheets.add(2)
