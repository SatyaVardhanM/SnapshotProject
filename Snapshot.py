from tkinter import *
from datetime import datetime
from PIL import ImageTk, Image,ImageGrab
import pyautogui as pg
import os
import xlsxwriter
from natsort import natsorted
import math
import time
import sys
import customtkinter
import logging
from threading import Thread

class _Log:
    def __init__(self):
        logging.basicConfig(filename="Temp.log",format='%(asctime)s %(message)s',filemode='w')
        streamHandler  =logging.StreamHandler()
        self.logger = logging.getLogger("Snapshot")
        self.logger.removeHandler(streamHandler)
        self.logger.setLevel(logging.DEBUG)
    def Debug(self,message):
        return self.logger.debug("DEBUG : "+message)
    def Error(self,message):
        return self.logger.error("ERROR : "+message)
    def Info(self,message):
        return self.logger.info("INFO : "+message)


class ScreenshotImageEditor():
    def __init__(self):
        self.master =None
        self.Save = None
        self.canvas = None
        self.brush_size = 8
        self.Markerimg = None
        self.imagePath = None
        self.back_img = None
        self.buttonClickMotion = None
        self.canvasImage_x1 = None
        self.canvasImage_y1 = None
        self.canvasImage_x2 = None
        self.canvasImage_y2 = None
        self.imageMaxWidth = None
        self.imageMaxHeight = None
        self.draw_layer = None
        self.canvas_image_id = None


    def SaveImageFile(self, *args):
        # Merges the layers and saves
        Image.alpha_composite(self.back_img, self.draw_layer).convert("RGB").save(self.imagePath)
        self.master.withdraw()

    def ResizeImageFile(self,ImageFileData,width,height):
        image = ImageFileData
        if image.width>width and image.height>height:
            image = image.resize((width,height))
        elif image.width >width:
            image = image.resize((width,image.height))
        elif image.height >height:
            image =image.resize((image.width,height))
        else :
            temp_width = width - image.width
            temp_height = height - image.height
            if(temp_width<temp_height):
                image = image.resize((image.width+temp_width,image.height+temp_width))
            elif(temp_width>temp_height):
                image = image.resize((image.width+temp_height,image.height+temp_height))
            else:
                image = image.resize((image.width+temp_height,image.height+temp_height))
        return image

    def resizeImageAndCreateCanvas(self,path,imageData):
        self.master = Toplevel()
        self.master.title("Edit")
        self.Save = customtkinter.CTkButton(self.master,text="Save",command=self.SaveImageFile)
        self.Save.pack()
        self.canvas = customtkinter.CTkCanvas(self.master, width=self.master.winfo_screenwidth(),height=self.master.winfo_screenheight())
        self.canvas.pack()
        self.imageMaxWidth = self.master.winfo_screenwidth()
        self.imageMaxHeight = self.master.winfo_screenheight()
        self.imagePath = path
        self.back_img = imageData.convert("RGBA") # Ensure RGBA
        self.draw_layer = Image.new("RGBA", self.back_img.size, (255, 255, 255, 0))
        self.last_x = None
        self.last_y = None
        

        img_width = self.back_img.width
        img_height = self.back_img.height
        isLargeImage = False
        if(self.master.winfo_screenwidth()-self.back_img.width<100):
            img_width = img_width-100
            isLargeImage =True
        if(self.master.winfo_screenheight()-self.back_img.height<100):
            img_height = img_height-100
            isLargeImage =True
        
        self.Markerimg  = ImageTk.PhotoImage(self.ResizeImageFile(self.back_img,img_width,img_height))
        
        if(isLargeImage):
            self.canvasImage_x1 = (self.master.winfo_screenwidth()-img_width)/2
            self.canvasImage_y1 = (self.master.winfo_screenheight()-img_height)/4
        else:
            self.master.geometry(f"{img_width+30}x{img_height+60}+0+0")
            self.canvasImage_x1 = 15
            self.canvasImage_y1 = 15

        self.canvas_image_id = self.canvas.create_image(self.canvasImage_x1, self.canvasImage_y1, anchor="nw", image=self.Markerimg)
        self.master.resizable(False, False)
        self.buttonClickMotion = self.canvas.bind("<B1-Motion>", self.draw_brush)
        self.canvas.bind("<ButtonRelease-1>", lambda e: self.reset_last_pos())

    def reset_last_pos(self):
        self.last_x, self.last_y = None, None

    def draw_brush(self, event):
        from PIL import ImageDraw
        draw = ImageDraw.Draw(self.draw_layer)
        r = self.brush_size
        color = (255, 255, 0, 80)

        # 1. Calculate local coordinates (where 0,0 is the corner of the image)
        curr_x = event.x - self.canvasImage_x1
        curr_y = event.y - self.canvasImage_y1

        # 2. Draw the smooth connection
        # We only draw if the mouse is actually inside the image bounds
        if 0 <= curr_x <= self.back_img.width and 0 <= curr_y <= self.back_img.height:
            if self.last_x is not None:
                # Use a line to connect last point to current point
                draw.line([self.last_x, self.last_y, curr_x, curr_y], fill=color, width=r*2)
            
            # Draw the circle at the current tip
            draw.ellipse([curr_x-r, curr_y-r, curr_x+r, curr_y+r], fill=color)

            # 3. Update tracking
            self.last_x, self.last_y = curr_x, curr_y

            # 4. Refresh display
            combined = Image.alpha_composite(self.back_img, self.draw_layer)
            self.Markerimg = ImageTk.PhotoImage(combined)
            self.canvas.itemconfig(self.canvas_image_id, image=self.Markerimg)

class Snapshot: 
    def __init__(self,_screenshotEditor,_log):
        #TextBoxes
        self.FilePath_textBox = None
        self.TestCase_textBox = None
        self.TestCaseImageTitle_textBox = None
        self.CaptureMode_textBox = None
        self.ExcelFileName_textBox = None
        #Labels
        self.FilePath_label = None
        self.TestCase_label = None
        self.TestCaseImageTitle_label =None
        self.CaptureMode_label =None
        self.ExcelFileName_label = None
        #Buttons
        self.Capture_button = None
        self.Excel_button = None
        self.Check_button = None
        #Switch
        self.switchMode =None
        #Event Binding
        self.motionBind = None
        self.buttonReleaseBind = None
        self.buttonClickBind = None
        ###
        self.isCaptureButtonClicked = False
        self.rectangle = None
        self.testCaseTitleDict = {}
        self.baseDirectory = None
        self.TestCase=None
        self.modeDropDownCall = None
        self.SelectedMode = None
        self.borderColor =None
        self.startPoint = None
        self.endPoint = None
        self.colorMode =None

        #Creating a main window model
        self.win = customtkinter.CTk()
        self.win.resizable(False, False)
        self.win.iconbitmap(os.getcwd()+'\\WindowIcon.ico')
        self.win.title("Snapshot") 
        self.winNew =None
        self.Error_label =None
        self._log = _log
        self.screenshotEditor = _screenshotEditor
        self.imageMaxWidth = self.win.winfo_screenwidth()
        self.imageMaxHeight = self.win.winfo_screenheight()

    def GetImageFileName(self):
        maxCount = 0
        newImgDirectory = self.baseDirectory +"TestCase_"+self.TestCase+"\\"
        if not os.path.exists(newImgDirectory):
            os.makedirs(newImgDirectory)

        for i in range(len(os.listdir(newImgDirectory)),0,-1):
            if os.path.isfile(os.path.join(newImgDirectory, f"Screenshot_{i}.png")):
                if(maxCount<i):
                    maxCount=i

        fileName = newImgDirectory+"Screenshot_"+str(maxCount+1)+".png"

        self.testCaseTitleDict[fileName] = self.TestCaseImageTitle_textBox.get().replace("\n", "")
        self.TestCaseImageTitle_textBox.delete(0,END)
        return fileName

    def EditScreenshot(self,screenshotImage):
        self.screenshotEditor.resizeImageAndCreateCanvas(self.GetImageFileName(), screenshotImage)
        self.win.deiconify()
        
    ################################    
    #Grab and Take Screenshot
    def GrabScreenshot(self,event):
        self.winNew.wm_attributes("-transparent","green")
        self.rectangle = self.canvas.create_rectangle(0,0,0,0,width=3,fill="green",outline = "red")
        self.startPoint = [int(event.x),int(event.y)]
        self.motionBind =self.winNew.bind('<Motion>',self.ReleaseScreenshot)

    def ReleaseScreenshot(self,event):
        self.canvas.coords(self.rectangle,self.startPoint[0],self.startPoint[1],int(event.x),int(event.y))
        self.canvas.pack()
        self.buttonReleaseBind =self.winNew.bind('<ButtonRelease-1>',self.GetEndPoint)

    def GetEndPoint(self,event):
        self.endPoint= [int(event.x),int(event.y)]
        self.RemoveEventsAndCloseCaptureWindow()
        self.EditScreenshot(self.Screenshot(self.startPoint,self.endPoint))

    def RemoveEventsAndCloseCaptureWindow(self):
        self.canvas.delete(self.rectangle)
        self.winNew.unbind("<Button-1>",self.buttonClickBind)
        self.winNew.unbind("<Motion>",self.motionBind)
        self.winNew.unbind("<ButtonRelease-1>",self.buttonReleaseBind)
        self.winNew.destroy()
    ##############################

    def OnCapture(self):
        self.baseDirectory = (self.FilePath_textBox.get().replace("\n", "").replace(" ", "")+"\\").replace("\\\\", "\\")
        self.TestCase =self.TestCase_textBox.get().replace(" ", "").replace("\n", "")
        self.isCaptureButtonClicked =True
        errorMessage = self.ValidateInputData(self.baseDirectory,self.TestCase)
        selectedMode = self.modeDropDownCall.get()
        try:
            if(len(errorMessage)==0 and self.Error_label !=None):
                self.Error_label.configure(text ='')
                self.TestCase_textBox.configure(border_color=self.borderColor)
                self.FilePath_textBox.configure(border_color=self.borderColor)

                if(selectedMode=="Full"):
                    self.ScreenSnipper(isFullscreen = True)
                elif(selectedMode=="Selected"):
                    self.ScreenSnipper()
        except Exception as e:
            self._log.Error(f"Error Occured while taking a screenshot. Error:{sys.exc_info()[1]}")
            self.winNew.destroy()
            self.win.deiconify()
        
    def ScreenSnipper(self,isFullscreen = False):
        if(isFullscreen):
            self.win.withdraw()
            time.sleep(0.5)
            self.EditScreenshot(self.Screenshot([0,0],[self.win.winfo_screenwidth(),self.win.winfo_screenheight()]))
            time.sleep(1)
            self.win.deiconify()
        else:
            self.winNew =Toplevel()
            self.winNew.withdraw()
            self.winNew.geometry(str(self.winNew.winfo_screenwidth())+"x"+str(self.winNew.winfo_screenheight()))
            self.winNew.wm_attributes('-fullscreen','True')
            self.winNew.wm_attributes('-alpha',0.5)
            self.canvas = customtkinter.CTkCanvas(self.winNew,width=int(self.winNew.winfo_screenwidth()),height=int(self.winNew.winfo_screenheight()))
            self.winNew.resizable(False, False)
            self._log.Info("ScreenshotMode:Selected, window Created.")
            self.winNew.deiconify()
            self.buttonClickBind =self.winNew.bind('<Button-1>',self.GrabScreenshot)
            self._log.Error("ScreenshotMode:Selected, window closed.")
            self.win.withdraw()


    def Screenshot(self,initialPosition,finalPosition):
        value=()
        x1,y1,x2,y2= abs(initialPosition[0]),abs(initialPosition[1]),abs(finalPosition[0]),abs(finalPosition[1]) 
        if x1!=x2 and y1!=y2:
            if(x2<x1):
                if(y2<y1):
                    value=(x2,y2,x1,y1)
                if(y2>y1):
                    value=(x2,y1,x1,y2)
            else:
                if(y2<y1):
                    value=(x1,y2,x2,y1)
                if(y2>y1):
                    value=(x1,y1,x2,y2)
        return ImageGrab.grab(bbox=value)
        
    def ValidateInputData(self,filePathField,testCaseField=None):
        errorMessage = ""
        
        self.FilePath_textBox.configure(border_color=self.borderColor)
        self.TestCase_textBox.configure(border_color=self.borderColor)
        if(filePathField=="" or filePathField=="\\"):
            errorMessage +="Folder Path cannot be blank."
            self.FilePath_textBox.configure(border_color='red')
        elif os.path.exists(filePathField) is False and  os.access(filePathField, os.R_OK|os.W_OK|os.X_OK) is False:
                errorMessage +="Invalid Folder path. i.e. Is not Accessible OR Does not exist."
                self.FilePath_textBox.configure(border_color='red')
                
        if(self.isCaptureButtonClicked):
            if(len(testCaseField)<1):
                errorMessage +="Test Case value cannot be blank."
                self.TestCase_textBox.configure(border_color='red')
            elif(testCaseField.isnumeric()):
                if(int(testCaseField)<1):
                    errorMessage +="Test case value must be grater then 0(Zero)."
                    self.TestCase_textBox.configure(border_color='red')
            else:
                errorMessage +="Test case value must be numeric."
                self.TestCase_textBox.configure(border_color='red')

        if(len(errorMessage)>0):
            self._log.Error(errorMessage)
            self.Error_label.configure(text=errorMessage,text_color='red')

        self.isCaptureButtonClicked=False
        return errorMessage



    def GenerateExcel(self):
        self.Error_label.configure(text="")
        self.Capture_button.configure(state = 'disabled')
        self.win.config(cursor='watch')
        validation_diectory = self.FilePath_textBox.get().replace("\n", "").replace(" ", "")
        source_diectory = (validation_diectory+"\\").replace("\\\\", "\\")
        excelFileName =self.ExcelFileName_textBox.get().replace("\n", "").replace(" ", "")
        errorMessage = ""

        if(len(excelFileName)==0):
            excelFileName = 'Test_Evidence_Sheet'
        if(self.isCaptureButtonClicked is False):
                errorMessage =self.ValidateInputData(validation_diectory)
                
        if(len(errorMessage)==0):
            
            self._log.Info(f"Excel file creation started")
            workbook = xlsxwriter.Workbook(f'{source_diectory}{excelFileName}_{datetime.now().strftime("%d%m%Y_%H%M%S")}.xlsx')
            worksheet = workbook.add_worksheet()
            cell_format = workbook.add_format()
            cell_format.set_bold()
            image_row=0
            image_col =0
            errorWhileWritingFile = ''
            try:
                for directory in natsorted(os.listdir(source_diectory)):
                    directory_path = os.path.join(source_diectory,directory)
                    if(os.path.isdir(directory_path) is False):
                        continue
                    
                    worksheet.write(image_row,image_col,directory,cell_format)
                    
                    rowGap = 2
                    image_row+=2
                    for file in natsorted(os.listdir(directory_path)):
                        self._log.Info(f"Attaching file :  {file} in Excelfile started")
                        
                        imagePath = os.path.join(directory_path,file)
                        image = Image.open(imagePath)

                        if(self.testCaseTitleDict.get(imagePath)):
                            worksheet.write(image_row-1,image_col,self.testCaseTitleDict[imagePath],cell_format)
                            
                        
                        if(image.height>(self.imageMaxHeight/2) or image.width>(self.imageMaxWidth/2)):
                            x_scale = ((image.height*0.6)/self.imageMaxHeight)+0.6
                            y_scale = ((image.width*0.6)/self.imageMaxWidth)+0.6
                            if(x_scale>1):
                                x_scale = (1/x_scale)-0.2
                                y_scale = (1/x_scale)-0.2
                            if(y_scale>1):
                                x_scale = (1/y_scale)-0.2
                                y_scale = (1/y_scale)-0.2
                            rowSize = round((20*((image.height*y_scale)/(self.imageMaxHeight/2))))
                            worksheet.insert_image(image_row,image_col, imagePath,
                            {
                                'x_scale':x_scale,'y_scale':y_scale,
                                'positioning':1
                            })
                        else:
                            rowSize = math.ceil((20*((image.height)/(self.imageMaxHeight/2))))
                            worksheet.insert_image(image_row,image_col, imagePath,
                            {
                                'positioning':1
                            })

                        
                        image_row+=rowSize+rowGap
                        self._log.Info(f"Attaching file :  {file} in Excelfile completed")   
                    
                    image_row+=rowGap
            except:
                errorWhileWritingFile = "Error Occured while writing Excel file"
                self._log.Error(errorWhileWritingFile)  
            workbook.close()
            if(errorWhileWritingFile==''):
                self.Error_label.configure(text="Excel Generated successfully",text_color='green')
            else:
                self.Error_label.configure(text=errorWhileWritingFile,text_color='red')
            self._log.Info(f"Excel file creation Completed")
            
        self.Capture_button.configure(state = 'normal')
        self.win.config(cursor='')
        

    def ChangeColorTheme(self):
        color = 'Light'
        self.borderColor = '#B2BEB5'
        val = self.colorMode.get()
        if(val==1):
            color='Dark'
            self.borderColor = 'gray35'
        
        self._log.Info(f"Application theme changed to {color} Mode")
        customtkinter.set_appearance_mode(color)  # Modes: "System" (standard), "Dark", "Light"
    

    def CreateThreadForExcel(self):
        Thread(target=self.GenerateExcel).start()
    

    def CreateAndInitializeForm(self):
        self._log.Info(f"Application Started")
        try:
            customtkinter.set_appearance_mode('Light')
            customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
            self.borderColor = '#B2BEB5'#Light Gray color for borders
            self.colorMode = customtkinter.IntVar()
            self.switchMode = customtkinter.CTkSwitch(self.win,text='DarkMode',variable=self.colorMode,command=self.ChangeColorTheme)

            self.FilePath_label = customtkinter.CTkLabel(self.win,text="File Path",font=('Public Sans',12, 'bold'))
            self.FilePath_textBox = customtkinter.CTkEntry(self.win,width=300,border_color=self.borderColor)
            self.TestCaseImageTitle_label = customtkinter.CTkLabel(self.win,text="Image Title",font=('Public Sans',12, 'bold'))
            self.TestCaseImageTitle_textBox = customtkinter.CTkEntry(self.win,width=300,border_color=self.borderColor)
            self.ExcelFileName_label = customtkinter.CTkLabel(self.win,text="Excel FileName",font=('Public Sans',12, 'bold'))
            self.ExcelFileName_textBox = customtkinter.CTkEntry(self.win,width=300,border_color=self.borderColor)
            self.TestCase_label = customtkinter.CTkLabel(self.win,text="TestCase No.",font=('Public Sans',12, 'bold'))
            self.TestCase_textBox = customtkinter.CTkEntry(self.win,border_color=self.borderColor,width=100)
            self.modeDropDownCall = customtkinter.StringVar()
            self.modeDropDownCall.set("Selected")
            self.CaptureMode_label = customtkinter.CTkLabel(self.win,text="Mode",font=('Public Sans',12, 'bold'))
            self.CaptureMode_textBox = customtkinter.CTkOptionMenu(self.win,values=["Selected","Full"],variable=self.modeDropDownCall,font=('Public Sans',12, 'bold'))
            
            self.Error_label =customtkinter.CTkLabel(self.win,text='',font=('Public Sans',12, 'bold'))
            
            self.Excel_button = customtkinter.CTkButton(self.win,text="GENERATE EXCEL",font=('Public Sans',12, 'bold'),command= self.CreateThreadForExcel,hover_color='darkgreen',fg_color='green')
            self.Capture_button = customtkinter.CTkButton(self.win,text="CAPTURE",font=('Public Sans',12, 'bold'),command=self.OnCapture,height=120)

            self.switchMode.grid(row=0,column=2,columnspan=2)
            self.FilePath_label.grid(row=1,column=0,sticky=W,pady=1,padx=5)
            self.FilePath_textBox.grid(row=1,column=1,pady=1,padx=2,columnspan = 4)
            self.TestCase_label.grid(row=2,column=0,sticky=W,padx=5)
            self.TestCase_textBox.grid(row=2,column=1,padx=2)
            self.CaptureMode_label.grid(row=2,column=2,sticky=W,padx=5)
            self.CaptureMode_textBox.grid(row=2,column=3,padx=2,ipadx=5)
            self.TestCaseImageTitle_label.grid(row=3,column=0,sticky=W,pady=1,padx=5)
            self.TestCaseImageTitle_textBox.grid(row=3,column=1,columnspan = 4,pady=1,padx=2)
            self.ExcelFileName_label.grid(row=4,column=0,sticky=W,pady=1,padx=5)
            self.ExcelFileName_textBox.grid(row=4,column=1,columnspan = 4,pady=1,padx=2)
            self.Capture_button.grid(row = 1, column = 5, columnspan = 2,rowspan = 3,padx=10,pady=10)
            self.Excel_button.grid(row=4,column=5,columnspan = 2,pady=10,padx=10)
            self.Error_label.grid(row=6,column=0,columnspan = 7)
            
        except:
            self._log.Debug(f"Application is closed due to fatal error. ERROR:{sys.exc_info()[1]}" )
        finally:
            self.win.mainloop()




if __name__ == '__main__':
    app = Snapshot(ScreenshotImageEditor(),_Log())
    app.CreateAndInitializeForm()
