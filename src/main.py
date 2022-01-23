from tkinter import filedialog, messagebox
from tkinter import *
from presentationMaker import makePresentation
from ScrollFrame import ScrollFrame
from OrderedSet import OrderedSet
import os


class Slide:
    def __init__(self):
        self.nameVar = StringVar()


class App:
    def __init__(self):
        self.tk = Tk()
        self.tk.title("PresentIt")
        self.dialog = None

        self.header = Frame(self.tk)
        self.header.grid(row=0, column=0)
        self.row = 0
        Label(self.header, text="Workspace:").grid(row=0, column=0, sticky="W")
        self.workspace = ""
        self.workspaceVar = StringVar()
        self.workspaceVar.set("Please select a workspace")
        self.workspaceLabel = Label(self.header, textvariable=self.workspaceVar)
        self.foregroundColor = self.workspaceLabel["fg"]
        self.workspaceLabel["fg"] = "gray"
        self.workspaceLabel.grid(row=0, column=1)
        button = Button(self.header, text="Select", command=self.selectWorkspace)
        button.grid(row=0, column=2)
        self.addRow()

        self.name = self.field("Name:")
        self.subtitle = self.field("Subtitle:")
        self.backgroundImage = self.field("Background image:")

        Label(self.header, text="Image extension:").grid(row=self.row, column=0, sticky="W")
        self.imageExtension = StringVar()
        self.imageExtension.set("PNG")
        menu = OptionMenu(self.header, self.imageExtension, "BMP", "TIFF", "PNG", "JPEG", "JPG", "GIF")
        menu.grid(row=self.row, column=1, sticky="W")
        self.addRow()

        Button(self.header, text="New slide", command=self.newSlide).grid(row=self.row, column=0, sticky="W")
        self.addRow()
        Label(self.header, text="Slides:").grid(row=self.row, column=0, sticky="W")
        self.addRow()
        self.slideList = ScrollFrame(self.tk, width=280, height=250)
        self.slideList.grid(row=1, column=0, sticky="W")
        self.addRow()
        self.slides = OrderedSet()

        button2 = Button(self.tk, text="Create presentation", command=self.createPresentation)
        button2.grid(row=2, column=0, sticky="W")

    def field(self, name: str) -> StringVar:
        Label(self.header, text=name).grid(row=self.row, column=0, sticky="W")
        var = StringVar()
        Entry(self.header, textvariable=var).grid(row=self.row, column=1, sticky="W")
        self.addRow()
        return var

    def addRow(self):
        self.row += 1

    def newSlide(self):
        def delete():
            nonlocal slideFrame, slide
            slideFrame.destroy()
            self.slides.remove(slide)

        slide = Slide()
        self.slides.add(slide)
        slideFrame = Frame(self.slideList.viewPort)
        slideFrame.pack()
        Label(slideFrame, text="Name:").grid(row=0, column=0)
        Entry(slideFrame, textvariable=slide.nameVar).grid(row=0, column=1)
        Button(slideFrame, text="-", command=delete).grid(row=0, column=2)

    def createPresentation(self):
        def callback(workspace):
            if not self.name.get():
                messagebox.showinfo(
                    "Please choose a name for this presentation",
                    "You need to choose a name to create a presentation"
                )
                return
            os.chdir(workspace)
            name = f"{self.name.get()}.pptx"
            titleSlide = {
                "type": "title",
                "title": self.name.get()
            }

            subtitle = self.subtitle.get()
            if subtitle:
                titleSlide["subtitle"] = subtitle

            backgroundImage = self.backgroundImage.get()
            if backgroundImage:
                titleSlide["backgroundImage"] = backgroundImage

            slides = [titleSlide]

            for slide in self.slides:
                slideName = slide.nameVar.get()
                slideDesc = {
                    "type": "content",
                    "title": slideName,
                    "text": f"{slideName}.txt"
                }

                imagePath = f"{slideName}.{self.imageExtension.get()}"
                if os.path.exists(imagePath):
                    slideDesc["pictures"] = [
                        {
                            "file": imagePath,
                            "horizontalAlignment": "trailing"
                        }
                    ]
                slides.append(slideDesc)

            makePresentation({"name": name, "slides": slides})

        self.getWorkspace(callback)

    def selectWorkspace(self):
        self.workspace = filedialog.askdirectory(title="Select presentation workspace")
        if self.workspace:
            self.workspaceVar.set(self.workspace)
            self.workspaceLabel["fg"] = self.foregroundColor
        else:
            self.workspaceVar.set("Please select a workspace")
            self.workspaceLabel["fg"] = "gray"

    def getWorkspace(self, callback):
        def okPressed():
            self.selectWorkspace()
            self.getWorkspace(callback)

        if self.workspace:
            if self.dialog:
                self.hideDialog()
            callback(self.workspace)
        elif not self.dialog:
            self.showMessage(callback=okPressed)

    def showMessage(self, callback):
        self.dialog = Toplevel()
        self.dialog.title("Please select a workspace")
        Label(self.dialog, text="You need a workspace to create a presentation").pack()
        buttonFrame = Frame(self.dialog)
        buttonFrame.pack()
        Button(buttonFrame, text="OK", command=callback).grid(row=0, column=0)
        Button(buttonFrame, text="Cancel", command=self.hideDialog).grid(row=0, column=1)

    def hideDialog(self):
        self.dialog.destroy()
        self.dialog = None

    def run(self):
        self.tk.mainloop()


app = App()
app.run()
