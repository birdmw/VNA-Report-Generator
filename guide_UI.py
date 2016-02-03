#!/usr/bin/python

# guide_UI.py
# Matthew Bird
# 2/3/2016

from Tkinter import *
from functools import partial
import tkFileDialog
import numpy as np
import pandas as pd

# DEPENDENCIES:
#  Download and install Anaconda for Windows Python 2.7 - https://www.continuum.io/downloads#_windows


class GUI(object):
    # A GUI object that uses TKinter to ask a user for relevant information to generate a guide.csv file
    def __init__(self):
        '''
        Declaration of internal variables, including widgets to be used for each window
        :return:
        '''
        self.root = Tk()
        self.root_labels = []
        self.root_entries = []
        self.root_buttons = []
        self.sNp_count = IntVar()
        self.plot_count = IntVar()

        self.window2 = None
        self.window2_labels = []
        self.window2_entries = []
        self.window2_buttons = []
        self.window2_radio_buttons = []
        self.window2_path_vars = []
        self.window2_port_vars = []

        self.window3 = None
        self.window3_labels = []
        self.window3_entries = []
        self.window3_buttons = []
        self.plot_param_choices = []
        self.p_bode_domain = []
        self.plot_ts_choices = []

    def root_window(self):
        '''
        This is the first window to appear, and asks users for how many plots and touchstone files
        :return:
        '''
        self.root_labels.append(Label(self.root, text="How many touchstone files? ").grid(row=1, column=1, sticky=E))
        self.root_entries.append(Entry(self.root, textvariable=self.sNp_count, width=3).grid(row=1, column=2, sticky=W))
        self.root_labels.append(Label(self.root, text="How many plots? ").grid(row=2, column=1, sticky=E))
        self.root_entries.append(Entry(self.root, textvariable=self.plot_count, width=3).grid(row=2, column=2,
                                                                                              sticky=W))
        self.root_buttons.append(Button(self.root, text="Ok", command=self.window2_window).grid(row=3, column=2))
        self.sNp_count.set(1)
        self.plot_count.set(1)
        mainloop()

    def window2_window(self):
        '''
        This is the second window asking use to pick out touchstone files
        :return:
        '''
        self.window2 = Toplevel()
        self.window2.lift()

        self.window2_labels.append(Label(self.window2, text="Choose").grid(row=0, column=0))
        self.window2_labels.append(Label(self.window2, text="Touchstone Files").grid(row=0, column=2))
        self.window2_labels.append(Label(self.window2, text="Port Map").grid(row=0, column=3))

        for i in xrange(int(self.sNp_count.get())):
            self.window2_path_vars.append(StringVar())
            self.window2_labels.append(Label(self.window2, text="TS "+str(i)).grid(row=i+1, column=0))
            self.window2_buttons.append(Button(self.window2, text='Browse...',
                                               command=partial(self.askopenfile, i)).grid(row=i+1, column=1))
            self.window2_entries.append(Entry(self.window2,
                                              textvariable=self.window2_path_vars[i]).grid(row=i+1, column=2))
            self.window2_port_vars.append(IntVar())
            self.window2_radio_buttons.append(Radiobutton(self.window2, text="1,3 to 2,4", value=0,
                                                          variable=self.window2_port_vars[i]).grid(row=i+1, column=3))
            self.window2_radio_buttons.append(Radiobutton(self.window2, text="1,2 to 3,4", value=1,
                                                          variable=self.window2_port_vars[i]).grid(row=i+1, column=4))
        self.window2_buttons.append(Button(self.window2, text='Ok',
                                           command=self.window3_window).grid(row=int(self.sNp_count.get())+1, column=4))

    def window3_window(self):
        '''
        This is the third window asking for Plot options
        :return:
        '''
        self.window3 = Toplevel()
        self.window3.lift()
        Label(self.window3, text="Plot Options").grid(row=0, column=0)
        Label(self.window3, text="Mag").grid(row=0, column=1)
        Label(self.window3, text="Phase").grid(row=0, column=2)
        Label(self.window3, text="sX_Y").grid(row=0, column=3+len(self.window2_path_vars))
        for ts_p_i in range(len(self.window2_path_vars)):
            Label(self.window3, text="TS" + str(ts_p_i)).grid(row=0, column=3+ts_p_i)
        for plt_i in range(int(self.plot_count.get())):
            Label(self.window3, text="Plot " + str(plt_i)).grid(row=plt_i+1, column=0)
        self.plot_ts_choices = [[0]*int(self.sNp_count.get()) for _ in xrange(int(self.plot_count.get()))]
        for i in range(len(self.plot_ts_choices)):
            for j in range(len(self.plot_ts_choices[i])):
                self.plot_ts_choices[i][j] = IntVar()
                self.plot_ts_choices[i][j].set(0)
        for plt_i in range(int(self.plot_count.get())):
            self.p_bode_domain.append(IntVar())
            Radiobutton(self.window3, text="Mag", value=0, variable=self.p_bode_domain[plt_i]).grid(row=plt_i+1, column=1)
            Radiobutton(self.window3, text="Phase", value=1, variable=self.p_bode_domain[plt_i]).grid(row=plt_i+1, column=2)
            for ts_i in range(int(self.sNp_count.get())):
                Checkbutton(self.window3, variable=self.plot_ts_choices[plt_i][ts_i]).grid(row=1+plt_i, column=3+ts_i)
                self.plot_ts_choices[plt_i][ts_i].set(0)
            self.plot_param_choices.append(StringVar())
            Entry(self.window3, textvariable=self.plot_param_choices[-1], width=8).grid(
                        row=plt_i+1, column=3+int(self.sNp_count.get()))
        Button(self.window3, text='Ok', command=self.close_gui).grid(row=int(self.plot_count.get())+1,
                                                                column=int(self.sNp_count.get()))
        self.window3.lift()

    def askopenfile(self, ts_number):
        '''
        This is the dialogue box that helps the user to navigate to sNp files
        :param ts_number:
        :return:
        '''
        self.window2_path_vars[ts_number].set(str(tkFileDialog.askopenfilename()))
        self.window2.lift()

    def close_gui(self):
        '''
        Destroys the GUI
        :return:
        '''
        self.root.destroy()

    def generate_guide(self):
        '''
        Creates a guide.csv file from the information gathered in the GUI
        :return:
        '''
        print "thing[plot][touchstone]"
        print "sNp count =", self.sNp_count.get()
        print "plot_count =", self.plot_count.get()
        print "paths =", [i.get() for i in self.window2_path_vars]
        print "port_mappings, [0, 1] =", [i.get() for i in self.window2_port_vars]
        print "plot parameters, [S11, S12, S21, S22] =", [i.get() for i in self.plot_param_choices]
        print "p_bode_domain:", [i.get() for i in self.p_bode_domain]
        print "plot_ts_choices:"
        for i in self.plot_ts_choices:
            print "for j in i:"
            for j in i:
                print j.get()
        guide = []
        column_names = ['', 'title', 'magphase']
        first_line = ['paths', '', '']
        second_line = ['portmap', '', '']
        for p in range(self.sNp_count.get()):
            column_names.append("sNp_"+str(p))
            first_line.append(str(self.window2_path_vars[p].get()))
            second_line.append('13_24' if self.window2_port_vars[p].get() == 0 else '12_34')
        guide.append(column_names)
        guide.append(first_line)
        guide.append(second_line)
        for p in range(self.plot_count.get()):
            row = ['plot_'+str(p), 'plot_'+str(p), 'm' if self.p_bode_domain[p].get() == 0 else 'p']+[str(self.plot_param_choices[p].get())]*self.sNp_count.get()
            guide.append(row)
        np_guide = np.array(guide)
        for row in np_guide:
            print row
        table = pd.DataFrame(np_guide[1:, 1:], index=np_guide[1:, 0], columns=np_guide[0, 1:])
        print table
        table.to_csv('guide.csv', sep=',')

    def run(self):
        '''
        Launches the GUI and then generates guide.csv from it
        :return:
        '''
        self.root_window()
        self.generate_guide()

if __name__ == "__main__":
    #  If this program is being invoked as main (as opposed to imported) it will create a GUI and run it
    g = GUI()
    g.run()
