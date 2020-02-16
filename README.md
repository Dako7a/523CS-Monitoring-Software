# 523CS-Monitoring-Software
This software is written for the purpose of independent use for the following purposes:
1. Monitoring conditions of the hardware & software we integrate for our customers. 
2. Efficiently reporting information directly to our surveillance department.
3. Organizing relevant information in such a way that is easily parsable.


### For Future Development
The [dev](https://github.com/Dako7a/523CS-Monitoring-Software/tree/master/523Tech-Camera-Surveillance-Monitor/dev) folder provides the necessary source code and configuration to freeze the application and self-contain the packages and libraries.
#### Required Libraries 
1. pip install wmi
2. pip install pywin32
3. pip install cx_freeze


### Required Software (Used for freezing the application)
1. Python (development has been done with *version 3.6+*) 
2. [Inno Compile Setup](http://www.jrsoftware.org/isinfo.php)

## How to Freeze and Deploy
1. Verify all credentials have been updated:
   * The.inno file (installer password requirements)
    * 523CamSurv.py (smtp information)
    * compile.py
2. Run the provided *compile.py* within the same directory as the desired *favicon.ico* and *523CamSurv.py* file
3. Run the .inno script to package the libraries and required documents
4. Package the *executable* with the *license.txt*, *favicon* and *readme.txt*

### Deployment Requirements
1. If installing onto a win7+ computer and Python is **NOT** installed, ensure you reach out to microsoft and grab the require VC_redist.x64 or VC_redist.x32. 

**523 Tech**
[Visit our website to learn more about how we can serve you!](http://www.523tech.com/)
