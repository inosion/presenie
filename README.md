# presenie

presenie is a PowerPoint Template renderer.

The Powerpoint is the template, and your data can be either YAML or JSON.
The resulting powerpoint can then be turned into a PDF, or share as it.

## Features

- Multipage support - produces new pages appropriately.
- Groups of objects
- Text replacement.
- Image replacement (from URLs in the JSON file)

PPTX Generator from `YAML` or `JSON` data

## Building & Running

**Create a fat executable jar**
* Build it
    ```
    gradle jar
    chmod +x ./presenie/build/libs/presenie.jar
    ```
* This can now be run `./presenie/build/libs/presenie.jar`

## Testing
```
gradle run --args="list -t assets/test/source.pptx"
gradle run --args="merge -t samples/sample_template2.pptx  -d samples/data.json -o ../output_$$.pptx"
```

`presenit.run merge -t docmaker/samples/sample_template2.pptx  -d docmaker/samples/data.json -o out/output_${RANDOM}.pptx`
