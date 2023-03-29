Aplication *add_file_to_database_v0.1*

## Description

---

The code is intended to read some specific files from experimental measurements of mass flow rate through some device. The measurements equipment allowed to acquire  
pressure, temperature, LPM (liter per minute) and SLPM (standard liter per minute) by means of two mass flow meters and three additional pressure points using pressure sensors. 
Finally, the mass flow rate is calculated based on delivered data.


**How does it work?**

A single measurement point consist of multiple samples e.g. 100, which are stored in one file. Several points can be measured within single experiment and all of them are stored in one folder.
During measurement campaign multiple configuration can be tested, thus the original measurement directory consist of numerous folders.

The advantage of the presented code is to gather all files into single database. It is more convenient way to extract data from files and allows to compare different configuration.

Beside of the application, the example database is shared. It is synthetic data, prepared only in order to demonstrate capabilities of the application.

*Operation*

In first step the application reads files from provided directory, calculate mass flow rate and average all data. Then, the raw data from files and new averaged values are copied to general database. 
Finally new single file with all data is saved in *.json file.

On current stage it is necessary to provide manually a directory of measurement files. What is more, one should fulfil date of conducted experiment and name of particular measurement set (*case name*). 
It is also useful to fill *description* and *additional info* sections. (See *Data to Read* section and *datadict* dictionary)

User can choose whether to create new database or read existing one from *.json file (see *read_database_from_file* function)

---

---
Function description:
- **final_add_files** - Create database from files and save it into *.json file. 
    - *Function arguments:* 
        - Dictionary with files directory,
        - Output JSON database directory.
- **display_tk_window** - Display window with database tree view. 
    - *Function arguments:* 
        - Temrorary database within the code,
        - (Optional) Window size (default 500x500).
- **db_quickview** - Quick preview of the database content. 
    - *Function argument:* 
        - Temporary database within the code.
- **display_avr_data_frame** - Display DataFrame from selected data from which a plot can be drawn. 
    - *Function arguments:* 
        - Temporary database within the code,
        - Date when the measurements files were created,
        - Measurement case name.*
- **plot_massflow_from_avr_data** - Draw a plot of mass flow rate in function of pressure .
    - *Function arguments:* 
        - Temporary database within the code,
        - Date when the measurements files were created,
        - Measurement case name,
        - (Optional) ==Plot size== (default 8x8).
- **read_database_from_file** - Read database from file.
    - *Function argument:*
        - Directory of a database
---