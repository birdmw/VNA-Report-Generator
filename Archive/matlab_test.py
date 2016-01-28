import matlab.engine
import matplotlib.pyplot as plt
import numpy as np

eng = matlab.engine.start_matlab()
touchstone_filename = 'default.S4P'
X, Y = 2, 1  #SddXY

eng.eval("touchstone_filename = '"+touchstone_filename+"';", nargout=0)
eng.workspace['X'] = X
eng.workspace['Y'] = Y

eng.eval("s_obj = sparameters(touchstone_filename);", nargout=0)
eng.eval("ghz = s_obj.Frequencies./1e9;", nargout=0)
eng.eval("s_raw = s2sdd(s_obj.Parameters,1);", nargout=0)
eng.eval("s_plot = squeeze(s_raw(X,Y,:));", nargout=0)
eng.eval("y_db = 20*log10(abs(s_plot));", nargout=0)
eng.eval("y_deg = (180./pi).*(angle(s_plot));", nargout=0)
eng.eval("Ps = s_obj.Parameters;", nargout=0)

db = np.array(eng.workspace['y_db']).T[0]
deg = np.array(eng.workspace['y_deg']).T[0]
ghz = np.array(eng.workspace['ghz']).T[0]

print db
print deg
print ghz
plt.plot(ghz, db)
plt.show()