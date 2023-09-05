Traceback (most recent call last):
  File ".\outageMasterScript.py", line 188, in <module>
    process_excel_files(args.input1,args.input2 ,args.output)
  File ".\outageMasterScript.py", line 176, in process_excel_files
    combined_data=fixes(combined_data,custom_rules)
  File ".\outageMasterScript.py", line 58, in fixes
    data = data.apply(apply_custom_classification_rules, custom_rules=custom_rules,axis=1)
  File "C:\Users\swx1283483\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\LocalCache\local-packages\Python37\site-packages\pandas\core\series.py", line 4357, in apply
    return SeriesApply(self, func, convert_dtype, args, kwargs).apply()
  File "C:\Users\swx1283483\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\LocalCache\local-packages\Python37\site-packages\pandas\core\apply.py", line 1043, in apply
    return self.apply_standard()
  File "C:\Users\swx1283483\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\LocalCache\local-packages\Python37\site-packages\pandas\core\apply.py", line 1101, in apply_standard
    convert=self.convert_dtype,
  File "pandas\_libs\lib.pyx", line 2859, in pandas._libs.lib.map_infer
  File "C:\Users\swx1283483\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\LocalCache\local-packages\Python37\site-packages\pandas\core\apply.py", line 131, in f
    return func(x, *args, **kwargs)
TypeError: apply_custom_classification_rules() got an unexpected keyword argument 'axis'
