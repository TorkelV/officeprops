# OfficeProps.js

A JavaScript library used to extract, edit or remove metadata in Microsoft Office and Open Office files.



#### Supports:
 * docx, dotx, docm, dotm
 * xlsx, xlsm, xlsb, xltm, xltx
 * pptx, ppsx, ppsm, pptm, potm, potx
 * ods, odt, odp

## Install:
Officeprops relies on [JSZip](https://stuk.github.io/jszip/) which has be included alongside this package if not using node.

Node:
```
npm install officeprops
```
Or include via cdn:
```
<script src='https://cdn.jsdelivr.net/npm/jszip@3.1.5/dist/jszip.min.js'></script>
<script src='https://cdn.jsdelivr.net/npm/officeprops@1.1.5/src/officeprops.js'></script>
```



## Usage:

The package adds a global "OFFICEPROPS" variable.

All functions take a [File](https://developer.mozilla.org/en-US/docs/Web/API/File) or [Blob](https://developer.mozilla.org/en-US/docs/Web/API/Blob) as parameter, or a [Buffer](https://nodejs.org/api/buffer.html#buffer_class_buffer) in Node.


##### Get metadata:
```javascript
OFFICEPROPS.getData(file).then(function(metadata){
    console.log(metadata.editable);
    console.log(metadata.readOnly);
}
```

##### Metadata Format:
Open office property names are translated to MS Office names. e.g. "editing-duration" becomes "totalTime".  
Returns the actual value, as well as a translated one for each property.
```javascript
metadata = {
    editable: {
        totalTime: {
            value: "PT3M43S", //actual value
            tvalue: "3 minutes", //translated value
            xmlPath: "office:meta/meta:editing-duration"
        },
        creator: {
            value: "Torkel Velure",
            tvalue: "Torkel Velure",
            xmlPath: "office:meta/meta:initial-creator"
        }
        //...see OFFICEPROPS.properties for full list of properties
    },
    readOnly: {
        slideTitles: ["Slide1, slide2, slide3"],
        titles: ["title1", "title2"]
        worksheets: ["sheet1"]
    }
}
```

##### Edit metadata:

```javascript
OFFICEPROPS.getData(file).then(function(metadata){
    metadata.editable.creator.value = "New author";
    OFFICEPROPS.editData(file,metadata).then(function(officeFile){
        console.log(officeFile) // blob/nodestream containing edited file.
    }
}
```


##### Remove metadata:
```javascript
OFFICEPROPS.removeData(file).then(function(officeFile){
    console.log(officeFile) // blob/nodestream with metadata removed.
}
```

For more, see [Examples](https://github.com/TorkelV/officeprops/blob/master/src/example/index.html) or [Tests](https://github.com/TorkelV/officeprops/blob/master/src/test/officeprops.test.js)

