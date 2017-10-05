# zipcode-distance-excel
This is a command line utility to automatically calculate the distance between two zipcodes and then put the results in an excel (.xlsx) file. It works for US postal codes only.

It was developed to help a colleague and is very application-specific.

##Prerequisites
Downloading is easy git, which is already on most systems, on ubuntu use:
 
```
sudo apt-get install git
```
For everyone else:

* [Git for Mac](https://git-scm.com/download/mac)

* [Git for Windows](https://git-scm.com/download/win)

## Installing

A step by step series of examples that tell you have to get a development env running

Say what the step will be

```
git clone https://github.com/esologic/zipcode-distance-excel
cd zipcode-distance-excel
python3 setup.py install
``` 

## Usage

in a directory with the .xlsx file that you want to modify, run:

```
python3 zde.py
```

The program skips the first row in the spreadsheet to avoid headers.

### Example Usage

Before:

![Before](https://user-images.githubusercontent.com/3516293/31227931-9ec99c96-a9a9-11e7-8d2d-e3c441878e83.PNG)

```
Which file would you like to modify?
[0] -  International Addresses.xlsx
[1] -  testbook.xlsx
File Number: 1
You've selected [testbook.xlsx] to edit.
Which sheet would you like to modify?
[0] -  Sheet1
[1] -  Sheet2
[2] -  Sheet3
Sheet Number: 0
You've selected [Sheet1] to edit.
Which column to read? (ie. A, B, AA): A
Which column to write result? (ie. A, B, AA): B
Point A Zipcode? 02114
Job Complete. 10 modifications made.
```
![After](https://user-images.githubusercontent.com/3516293/31227930-9eaf3946-a9a9-11e7-9384-d096fdfbc073.PNG)

## Authors

* **Devon Bray** - [site](http://www.esologic.com)
* **Miranda Lawell** - [site](https://www.mirandalawell.com)

## License

This project is licensed under the MIT License
