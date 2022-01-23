import tkinter as tk


# ************************
# Scrollable Frame Class
#
# See https://gist.github.com/mp035/9f2027c3ef9172264532fcd6262f3b01#file-tk_scroll_demo-py
# ************************
class ScrollFrame(tk.Frame):
    def __init__(self, parent, height=400, width=600):
        # create a frame (self)
        super().__init__(parent)

        # place canvas on self
        self.canvas = tk.Canvas(self, borderwidth=0, height=height, width=width)

        # place a frame on the canvas, this frame will hold the child widgets
        self.viewPort = tk.Frame(self.canvas)

        # place a scrollbar on self
        self.vsb = tk.Scrollbar(self, orient='vertical', command=self.canvas.yview)

        # attach scrollbar action to scroll of canvas
        self.canvas.configure(yscrollcommand=self.vsb.set)

        # pack scrollbar to right of self
        self.vsb.pack(side='right', fill='y')

        # pack canvas to left of self and expand to fill
        self.canvas.pack(side='left', fill='both', expand=True)

        # add view port frame to canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.viewPort, anchor='nw', tags='self.viewPort')

        # bind an event whenever the size of the viewPort frame changes.
        self.viewPort.bind('<Configure>', self.on_frame_configure)

        # bind an event whenever the size of the viewPort frame changes.
        self.canvas.bind('<Configure>', self.on_canvas_configure)

        # perform an initial stretch on render, otherwise the scroll region has a tiny border until the first resize
        self.on_frame_configure(None)

    def on_frame_configure(self, event):
        """ Reset the scroll region to encompass the inner frame """

        # whenever the size of the frame changes, alter the scroll region respectively.
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def on_canvas_configure(self, event):
        """ Reset the canvas window to encompass inner frame when required """
        canvas_width = event.width

        # whenever the size of the canvas changes alter the window region respectively.
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
