import win32com.client as win32
import datetime
import string, os
full_save_dir = ''

class set_CATIA_workbench_env:
    def __init__(self):
        # self.catapp = win32.Dispatch("CATIA.Application")
        self.env_name = {'Part_Design': 'PrtCfg', 'Product_Assembly': 'Assembly', 'Drafting': 'Drw'}
        self.catapp = win32.Dispatch("CATIA.Application")
        # add some CATIA-specific settings here (Seeing CATIA Automation manual-Application section)
        self.catapp.DisplayFileAlerts = False
        self.catapp.Visible = True

    def Part_Design(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Part_Design.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return

    def Product_Assembly(self):
        self.catapp.Visible = True
        self.catapp.StartWorkbench(self.env_name[self.Product_Assembly.__name__])
        try:
            temp = self.catapp.ActiveDocument
            temp.close()
        except:
            pass
        return

def File_open(target, dir):
    # 連結CATIA
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents
    # 將路徑設為目錄的文字宣告
    directory = str(dir)
    # directory = '\\'.join(directory.split('/'))
    print(directory)
    # gvar.folderdir = directory
    # 定義零件檔檔名
    part_dir = directory + target + '.CATPart'
    print(part_dir)
    # partdoc = document.Open("%s%s.%s" % (directory,target,"CATPart"))
    # 開啟該零件檔
    partdoc = document.Open(part_dir)
    return target + '.CATPart'
def Param_change(target, value):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    part = partdoc.Part
    parameter = part.Parameters
    # 按照介面輸入的參數找出相對應的面建出板子
    length = parameter.Item(target)
    realParam = parameter.Item(target)
    if target == "Length":
        length.Value = value
    elif target == "Width":
        length.Value = value
    elif target == "Height":
        length.Value = value
    elif target == "r":
        length.Value = value

    if target == "Shelf_Plane":
        realParam.Value = value

    # elif target == "Height":
    #     length.Value = value
    part.Update()
def open_assembly():
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents
    productdoc = document.Add("Product")
    product = productdoc.Product
    products = product.Products
def assembly_open_file(folder, target, type):
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    products = product.Products
    # print(type(gvar.folderdir))
    # print(type(target))
    # directory = '\\'.join(folder.split('/'))
    # 開啟 0為零件檔/1為組立檔 進入該組立檔
    if type == 0:
        filedir = "%s\%s.%s" % (folder, target, "CATPart")
    elif type == 1:
        filedir = "%s\%s.%s" % (folder, target, "CATProduct")
    print(filedir)
    import_file = filedir,
    list(import_file)
    productsvarient = products.AddComponentsFromFiles(import_file, "All")
    return productsvarient

def add_offset_assembly(element1, element2, dist, relation):
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    ref1 = product.CreateReferenceFromName("Product1/%s.1/!%s" % (element1, relation))
    ref2 = product.CreateReferenceFromName("Product1/%s.1/!%s" % (element2, relation))
    # 1表示偏移拘束
    constraint = constraints.AddBiEltCst(1, ref1, ref2)
    length = constraint.Dimension
    length.value = dist
    constraint.Orientation = 0
    product.Update()
    return True

def add_comp_offset_assembly(element1, element2, assm_element1, assm_element2, dist, spec, type):
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    if type == 0:
        ref1 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/!%s" % (element1, assm_element1))
        ref2 = product.CreateReferenceFromName(
            "Product1/%s/Assembly_Reference_%s/!PartBody/!%s" % (element2, spec, assm_element2))
    elif type == 1:
        ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName(
            "Product1/%s.1/Assembly_Reference_%s/!PartBody/!%s" % (element1, spec, assm_element1))
    elif type == 2:
        ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s.1/!Body_Panel/%s" % (element1, assm_element1))
    elif type == 3:
        ref1 = product.CreateReferenceFromName("Product1/%s/!rail/%s" % (element1, assm_element1))
        ref2 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/%s" % (element2, assm_element2))
    elif type == 4:
        ref1 = product.CreateReferenceFromName("Product1/%s/!rail/%s" % (element1, assm_element1))
        ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element2, assm_element2))
    if type == 5:
        ref1 = product.CreateReferenceFromName("Product1/%s.2/!PartBody/!%s" % (element1, assm_element1))
        ref2 = product.CreateReferenceFromName(
            "Product1/%s/Assembly_Reference_%s/!PartBody/!%s" % (element2, spec, assm_element2))
    constraint = constraints.AddBiEltCst(1, ref1, ref2)
    length = constraint.Dimension
    length.Value = dist
    if 'OPP' in spec:
        constraint.Orientation = 1
    else:
        constraint.Orientation = 0
    product.Update()
    return True

def saveas(save_dir, target, data_type):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    try:
        saveas = document.Item('%s%s' % (target, data_type))
        saveas.SaveAs('%s%s%s' % (save_dir, target, data_type))
    except:
        saveas = catapp.ActiveDocument
        saveas.SaveAs('%s%s%s' % (save_dir, target, data_type))
    finally:
        saveas.Save()

def save_dir(save_dir):
    time_now = datetime.datetime.now()
    # product_name = ('%s%s%s' % (int(width) // 10, int(height) // 10, int(depth) // 10))
    year = str((int(time_now.strftime('%Y')) % 1000) % 100)
    code_E = list(string.ascii_uppercase)
    month = code_E[int(time_now.strftime('%m')) - 1]
    code_e = list(string.ascii_lowercase)
    day = time_now.strftime('%d')
    hour = code_e[int(time_now.strftime('%H'))]
    minute = time_now.strftime('%M')

    file_name = ('%s%s%s%s%s' % (year, month, day, hour, minute))
    try:
        save_dir = '\\'.join(save_dir.split('/'))  # if using GUI to set file_dir
    except:  # if using API call method, which file_dir has benn processed
        pass
    newpath = os.path.join(save_dir, file_name)
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    return newpath
def saveas_specify_target(save_dir, target, data_type):
    catapp = win32.Dispatch('CATIA.Application')
    doc = catapp.Documents
    saveas = doc.Item('%s.%s' % (target, data_type))
    saveas.Save()
    saveas.Close()
def add_coincident_assembly(element1, element2, assm_element1, assm_element2, spec, type):
    catapp = win32.Dispatch("CATIA.Application")
    productdoc = catapp.ActiveDocument
    product = productdoc.Product
    products = product.Products
    # product = product.ReferenceProduct
    constraints = product.Connections("CATIAConstraints")
    if type == 0:
        ref1 = product.CreateReferenceFromName(
            "Product1/%s.1/!PartBody/!%s" % (element1, assm_element1))  # select box reference for assembly
        ref2 = product.CreateReferenceFromName("Product1/%s/Assembly_Reference_%s/!PartBody/!%s" % (
            element2, spec, assm_element2))  # select componment reference for assembly
    elif type == 1:
        ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/!%s" % (element1, assm_element1))
    elif type == 2:
        ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/!%s" % (element1, assm_element1))
    elif type == 3:
        ref1 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/!%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/!%s" % (element1, assm_element1))
    elif type == 4:
        ref1 = product.CreateReferenceFromName("Product1/%s.1/!PartBody/%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element1, assm_element1))
    elif type == 5:
        ref1 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element2, assm_element2))
        ref2 = product.CreateReferenceFromName("Product1/%s/!PartBody/%s" % (element1, assm_element1))

    constraint = constraints.AddBiEltCst(2, ref1, ref2)
    product.Update()
    return True