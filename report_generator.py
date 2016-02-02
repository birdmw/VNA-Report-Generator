#!/usr/bin/python

# report_generator.py
# Matthew Bird
# 1/26/2016

import matlab.engine
import sys
import os
import pandas as pd
import docxtpl
import glob
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx.shared import Inches


class Generator(object):
    '''Generator Class - This object combines several files (info.csv, template.docx, guide.csv, and many .sNp files)'''

    def __init__(self):
        self.template_path = None
        self.guide_path = None
        self.info_path = None
        self.sNp_dir = None
        self.figure_string = None
        self.figures = None

        self.df_info = pd.DataFrame()
        self.d_info = {}
        self.tpl = None
        self.df_guide = pd.DataFrame()
        self.sNp_paths = None
        self.sNp = []

        self.ts_manager = None  # self.ts_manager.db[0][1][0] ====> DB, TSfile 0, S21
        self.plots = None
        self.stop_frequency = None
        self.start_frequency = None
        self.step_frequency = None

    def read_from_paths(self):
        self.read_guide()
        self.read_template()
        self.read_info()
        self.read_sNps()

    def read_template(self):
        file_path = self.template_path + "\\template.docx"
        if os.path.isfile(file_path):
            print "template.docx found at", file_path
            self.tpl = docxtpl.DocxTemplate(file_path)
        else:
            print "template.docx not found"

    def read_guide(self):
        if self.guide_path:
            file_path = str(self.guide_path) + "\\guide.csv"
            if os.path.isfile(file_path):
                print "guide.csv found at", file_path
                self.df_guide = pd.read_csv(file_path)
                self.set_ts_port_maps()
            else:
                print "guide.csv not found, generating one instead"

                self.generate_guide()
                self.df_guide.to_csv('guide.csv')
                self.df_guide = pd.read_csv(file_path)
                self.set_ts_port_maps()

    def generate_guide(self):
        self.get_sNp_paths()
        figures = self.get_guides()
        col_names = ['', 'title', 'magphase']
        for i_p in range(len(self.sNp_paths)):
            col_names.append('sNp_'+str(i_p))
        self.df_guide = pd.DataFrame(columns=col_names) #put in column names
        s = ['paths', '', '']
        for p in self.sNp_paths:
            s.append(p)
        df_s = pd.DataFrame([s], columns=col_names) # put in row 1 with paths
        self.df_guide=self.df_guide.append(df_s)
        port_map = ['portmap', '', '']
        for _ in self.sNp_paths:
            port_map.append('13_24')
        df_portmap = pd.DataFrame([port_map], columns=col_names)
        self.df_guide = self.df_guide.append(df_portmap)
        for i_pl in range(len(figures)):
            s = ["plot_"+str(i_pl), "plot_"+str(i_pl), self.figures[i_pl][0]]
            for _ in self.sNp_paths:
                #print self.figures[i_pl]
                s.append("s"+self.figures[i_pl][1]+"_"+self.figures[i_pl][2])
            self.df_guide = self.df_guide.append(pd.DataFrame([s], columns=col_names))
        self.df_guide = self.df_guide.set_index(self.df_guide.columns[0])
        return self.df_guide

    def get_sNp_paths(self):
        if os.path.isfile(self.info_path + "\\guide.csv"):
            self.sNp_paths = self.df_guide.iloc[0,3:].values.tolist()
            print self.df_guide.iloc[0,3:].values.tolist()
        else:
            self.sNp_paths = glob.glob(self.sNp_dir+"/*.s*p")
        return self.sNp_paths

    def get_guides(self):
        split_figures = self.figure_string.split(",")
        self.figures = []
        for i_f in range(len(split_figures)):
            parsed = list()
            parsed.append(split_figures[i_f][-1])
            split_list = split_figures[i_f].split("_")
            X = "".join(ch for ch in split_list[0] if ch in "0123456789")
            Y = "".join(ch for ch in split_list[1] if ch in "0123456789")
            parsed.append(X)
            parsed.append(Y)
            self.figures.append(parsed)
        print "FIGS++++++++++++"
        print self.figures
        return self.figures

    def read_info(self):
        file_path = self.info_path + "\\info.csv"
        if os.path.isfile(file_path):
            print "info.csv found at", file_path
            self.df_info = pd.read_csv(file_path)
            self.d_info = self.df_info.set_index('key').iloc[:,:1].T.to_dict('list')
            for k in self.d_info.iterkeys():
                self.d_info[k] = self.d_info[k][0]
            print self.d_info
        else:
            print "info.csv not found"

    def set_ts_port_maps(self):
        '''
        using self.df_guide,
        make a list of binary port maps
        example:[0, 1, 1, 0, 1]

        where 0 = 13_24 and 1 = 12_13
        '''
        self.ts_port_maps = self.df_guide.iloc[1,3:]
        for p in range(len(self.ts_port_maps)):
            print "pm:", self.ts_port_maps[p]
            if self.ts_port_maps[p] == "13_24":
                self.ts_port_maps[p] = 0
            else:
                self.ts_port_maps[p] = 1

    def read_sNps(self):
        self.sNp_paths = self.get_sNp_paths()
        self.set_ts_port_maps()
        self.ts_manager = TouchstoneManager(self.sNp_paths, self.ts_port_maps)

    def write_from_data(self):
        self.gen_plots()
        if not(self.tpl):
            self.tpl = docxtpl.DocxTemplate()
        else:
            self.d_info['StartFrequency'] = str(round(self.start_frequency*1000))+"MHz"
            self.d_info['StopFrequency'] = str(round(self.stop_frequency))+"GHz"
            self.d_info['StepFrequency'] = str(round(self.step_frequency*1000))+"MHz"

            if not(self.df_info is None):
                index = self.df_info['key'].tolist()
                dut_start_index = index.index('<DUT>')
                dut_end_index = index.index('</DUT>')
                port_start_index = index.index('<PORT>')
                port_end_index = index.index('</PORT>')

            dut_table = self.df_info.fillna('').iloc[dut_start_index+1:dut_end_index].as_matrix()
            port_table = self.df_info.fillna('').iloc[port_start_index+1:port_end_index].as_matrix()

            sd1 = self.tpl.new_subdoc()
            sd2 = self.tpl.new_subdoc()
            sd3 = self.tpl.new_subdoc()
            sd4 = self.tpl.new_subdoc()

            table1 = sd1.add_table(rows=len(dut_table), cols=len(dut_table[0]))
            table1.style = 'TableGrid'
            for row in range(len(dut_table)):
                for col in range(len(dut_table[0])):
                    table1.cell(row, col).text = dut_table[row][col]

            table2 = sd2.add_table(rows=len(port_table), cols=len(port_table[0]))
            table2.style = 'TableGrid'
            for row in range(len(port_table)):
                for col in range(len(port_table[0])):
                    table2.cell(row, col).text = port_table[row][col]

            table3 = sd3.add_table(rows=len(self.sNp_paths)+1, cols=2)
            table3.style = 'TableGrid'
            table3.cell(0, 1).text = "Trace Name"
            table3.cell(0, 0).text = "sNp file"

            for row in range(len(self.sNp_paths)):
                table3.cell(row+1, 1).text = "TS"+str(row)
                table3.cell(row+1, 0).text = self.ts_manager.file_titles[row]

            for i in range(self.count_plots()):
                sd4.add_picture('img'+str(i)+'.png', width=Inches(6.5))

            self.d_info['DUTIdentification'] = sd1
            self.d_info['PortConfiguration'] = sd2
            self.d_info['TestSummaryLegend'] = sd3
            self.d_info['TestSummary'] = sd4

            self.tpl.render(self.d_info)
            print "rendered"

        self.tpl.save("generated_doc.docx")

    def count_plots(self):
        print "# of plots is:", str(len(self.df_guide.index)-2)
        return len(self.df_guide.index)-2

    def gen_plots(self):
        plot_count = self.count_plots()
        sNp_count = len(self.sNp_paths)
        for p in range(plot_count):
            plt.clf()
            y_arr, labels = [], []
            self.plots = []
            XY_list = []
            for t in range(sNp_count):
                parameter = self.df_guide['sNp_'+str(t)].loc[2+p]
                magphase = self.df_guide['magphase'].loc[2+p]
                title = self.df_guide['title'].loc[2+p]
                if not(pd.isnull(parameter)): # if parameters is not blank
                    # self.ts_manager.db[0][1][0] ====> DB, TSfile 0, S21
                    split_list = parameter.split("_")
                    X = int("".join(ch for ch in split_list[0] if ch in "0123456789"))-1
                    Y = int("".join(ch for ch in split_list[1] if ch in "0123456789"))-1
                    XY_list.append([X, Y])
                    labels.append("TS"+str(t)+" S"+str(X+1)+str(Y+1))
                    x = np.array(self.ts_manager.ghz[t][X][Y])
                    if magphase == 'm':
                        y = np.array(self.ts_manager.db[t][X][Y])
                        plt.ylabel('Magnitude (dB)')
                    else:
                        y = np.array(self.ts_manager.deg[t][X][Y])
                        plt.ylabel('Phase (deg)')
                    y_arr.append(y)
            plots_arr = []
            if all(XY == XY_list[0] for XY in XY_list):
                plt.title('S'+str(X+1)+str(Y+1))
            else:
                plt.title('Mixed Parameters')
            plt.xlabel('Frequency (GHz)')
            for ya in range(len(y_arr)):
                plots_arr.append(plt.plot(x, y_arr[ya], label=labels[ya]))
            self.stop_frequency = x[-1]
            self.start_frequency = x[1]
            self.step_frequency = (self.stop_frequency - self.start_frequency)/float(len(x))
            plt.legend(loc='lower right')
            plt.savefig("img"+str(p))
            self.plots.append(plt)


class TouchstoneManager(object):
    ''' TouchstoneManager
      Access like this:
       tsm = TouchstronManager(["Filepath1","Filepath2",...], port_maps)

    self.ts_manager.db[0][1][0] ====> DB, TSfile 0, S21
    '''
    def __init__(self, ts_paths, port_maps):
        self.pm = port_maps
        self.db, self.deg, self.ghz = [], [], []
        self.eng = matlab.engine.start_matlab()
        self.file_titles = None
        for ts_p in range(len(ts_paths)):
            dbX_params = []
            degX_params = []
            ghzX_params = []
            filename = str(ts_paths[ts_p])
            print "file =", filename
            self.eng.eval("touchstone_filename = '"+filename+"';", nargout=0)
            self.eng.eval("s_obj = sparameters(touchstone_filename);", nargout=0)
            self.eng.eval("ghz = s_obj.Frequencies./1e9;", nargout=0)
            self.eng.eval("s_raw = s2sdd(s_obj.Parameters,"+str(self.pm[ts_p]+1)+");", nargout=0)
            for X in range(1, 3):
                dbY_params = []
                degY_params = []
                ghzY_params = []
                for Y in range(1, 3):
                    self.eng.eval("s_plot = squeeze(s_raw("+str(X)+","+str(Y)+",:));", nargout=0)
                    self.eng.eval("y_db = 20 * log10(abs(s_plot));", nargout=0)
                    self.eng.eval("y_deg = (180./pi).*(angle(s_plot));", nargout=0)
                    db = np.array(self.eng.workspace['y_db']).T[0]
                    deg = np.array(self.eng.workspace['y_deg']).T[0]
                    ghz = np.array(self.eng.workspace['ghz']).T[0]
                    dbY_params.append(db)
                    degY_params.append(deg)
                    ghzY_params.append(ghz)
                dbX_params.append(dbY_params)
                degX_params.append(degY_params)
                ghzX_params.append(ghzY_params)
            self.db.append(dbX_params)
            self.deg.append(degX_params)
            self.ghz.append(ghzX_params)
        self.eng.quit()
        self.file_titles = self.get_sNp_titles(ts_paths)

    def earliest_difference(self, list_of_strings, reverse = False):
        '''
        :param list_of_strings:
        :return index: earliest point where strings are not the same
        '''
        min_length = len(min(list_of_strings, key=len))
        for i in range(min_length):
            if reverse == False:
                ch = list_of_strings[0][i]
            else:
                ch = list_of_strings[0][-i-1]
            for str in list_of_strings:
                if reverse == False:
                    if ch != str[i]:
                        return i
                else:
                    if ch != str[-i-1]:
                        return -i

    def get_sNp_titles(self, string_list):
        first = self.earliest_difference(string_list, reverse=False)
        last = self.earliest_difference(string_list, reverse=True)
        names = [x[first:last] for x in string_list]
        return names


def main():
    g = Generator()
    if 'help' in sys.argv:
        print
        print "All parameters are optional, but the default path right now is", os.getcwd()
        print
        print "A short list of available commands:"
        print "  i, info.csv filepath"
        print "  t, template.docx filepath"
        print "  g, guide.csv filepath"
        print "  d, directory with .sNp files"
        print "  p, figures to plot ie: 'f S1_1m,S1_2p,S2_2m'"
        print "     m is for magnitude, p is for phase"
        print
        print "USAGE EXAMPLE:"
        print "  C:\>python report_generator.py i c:\path\\to\info t c:\path\\to\\template"
        print
    else:
        g.info_path = sys.argv[sys.argv.index('i')+1] \
            if ('i' in sys.argv and sys.argv.index('i') != len(sys.argv)-1) else os.getcwd()
        g.template_path = sys.argv[sys.argv.index('t')+1] \
            if ('t' in sys.argv and sys.argv.index('t') != len(sys.argv)-1) else os.getcwd()
        g.guide_path = sys.argv[sys.argv.index('g')+1] \
            if ('g' in sys.argv and sys.argv.index('g') != len(sys.argv)-1) else os.getcwd()
        g.sNp_dir = sys.argv[sys.argv.index('d')+1] \
            if ('d' in sys.argv and sys.argv.index('d') != len(sys.argv)-1) else os.getcwd()
        g.figure_string = sys.argv[sys.argv.index('p')+1] \
            if ('p' in sys.argv and sys.argv.index('p') != len(sys.argv)-1) else "s1_1m,s1_1p,s2_1m,s2_1p"
        print g.figure_string
        g.read_from_paths()
        g.write_from_data()

if __name__ == "__main__":
    main()
