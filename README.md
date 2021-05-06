# Installation

- Install Brother Printer Driver
- Copy template files (.lbx) to template folder in executable directory
- Run this application

## API

### Get Printers
GET Request:
```
localhost:60024/api/printers
```

Example Response:
```json
[
  {
    "name": "Brother PT-E550W",
    "online": true,
    "mediaId": 259,
    "mediaName": "12 mm",
    "tapeLength": 0
  }
]
```

### Get Templates

GET Request:
```
localhost:60024/api/templates
```

Example Response:
```json
[
  {
    "name": "cap-label",
    "description": null,
    "medianame": "18 mm",
    "width": 1024,
    "length": 2496
  },
  {
    "name": "ic-label",
    "description": null,
    "medianame": "18 mm",
    "width": 1024,
    "length": 2496
  },
  {
    "name": "ind-label",
    "description": null,
    "medianame": "18 mm",
    "width": 1024,
    "length": 2496
  },
  {
    "name": "res-label",
    "description": null,
    "medianame": "18 mm",
    "width": 1024,
    "length": 2496
  }
]
```


### Print Label

Prints a label. Expects JSON Body with fields
- count: number of labels to print
- printer: Printer name to print on
- template: Template name (.lbx filename without extension)
- fields: field values defined in the template

Label format in the Printer needs to match label format defined in the template

POST:
```
localhost:60024/api/print
```

Example Response:
```json
{
	"count": 1,
	"printer":"Brother PT-E550W",
	"template":"ic-label",
	"fields": 
	{
		 "name":"AKSCT/Z BLACK ",
		 "desc":"AKSCT Series Black 2 Position 2.54 mm Pitch Closed Housing Miniature Jumper",
		 "manufacturer":"ASSMANN WSW Components",
		 "mpn":"AKSCT/Z BLACK",
		 "footprint":"—"
    }
}
```

### Preview Label

JSON Body: The same JSON as in print, but without count field

POST:
```
localhost:60024/api/preview
```

Response: base64 encoded image
(to test, the response body can be pasted here to show the preview image: https://codebeautify.org/base64-to-image-converter)



