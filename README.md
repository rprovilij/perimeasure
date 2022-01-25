# Perimeasure
 
"Peripheral measurements"; A computer interaction monitoring (CIM) software developed in partial fulfillment for my human-technology interaction master's thesis project.
In this project, I explore whether an association exists between individuals' work engagement and their computer use patterns. 

https://research.tue.nl/en/studentTheses/work-engagement-and-computer-use-patterns

To use, rename the .exe file to include a participant's unique id number. For example:

```
perimeasure-IDNR.exe --> perimeasure-1234.exe
```
Running the program will create a database that is stored on the participant's device. The following data is collected in 1-minute intervals:
* **Number of keystrokes** (without identifying the keys themselves)
* **Number of clicks**
* **Number of scroll events**
* **Number of corrector keys pressed** (Backspace, Delete, CTRL-Z)
* **Between-key intervals**: the average elapsed time between two subsequent keystrokes was measured each minute. If subsequent keys were pressed after 5s or longer, the between-key interval was not included in the average.)
* **Within-key intervals**: when a key is pressed and released, two distinct events are registered (i.e., key-down and key-up). The elapsed time between these two events is recorded for every key. Keystroke duration is therefore measured and averaged per minute.)
* **Time spent using video conferencing application**
