# zipcode-distance-excel
This is a command line utility to automatically calculate the distance between two zipcodes and then put the results in an excel (.xlsx) file. It works for US postal codes only.

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

```
python3 zde.py [yourfile.xlsx] [zipcode A] [Column of zipcode B's] [Results column]
```

By default the program skips the first row in the spreadsheet

### Example Usage


## Authors

* **Devon Bray** - [esologic](http://www.esologic.com)
* **Miranda Lawell** - [site](https://www.mirandalawell.com)

## License

This project is licensed under the MIT License
