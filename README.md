# Reckon

A simple time estimation function.

## Getting Started

Simply import into your Workbook or Personal Workbook in /XLSTART!


Call inside your work loop

```
Application.StatusBar = EstimateTick(index,count)
```

Works VERY well with [ExcelProgress](https://github.com/groovyMysterioso/ExcelProgress)
```
ProgressBar xUpdateMeter, EstimateTick(indexNumber,whatever.Count), indexNumber
```

## Authors

* **James Pritts** - *Initial work* - [GroovyMysterioso](https://github.com/GroovyMysterioso)

