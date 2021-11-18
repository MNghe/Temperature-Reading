# Temperature Reading
 The process of getting real time temperature reading from a PLC, have the information go into MATLAB with a GUI that will visually display it and then export the data into an Excel file.

The first step is to connect the Arduino with the thermistor to your workstation. 

Once that is completed, install the support package on MATLAB so the data can be read in MATLAB. 

Then you just run either MATLAB code and it should run fine. The tempReading will do just a fixed 100 samples unless adjusted, while the GUI script will open an actual GUI that would allow you to collect as much as you want and export that data.
