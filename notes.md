This might be the way we clone a sheet

Using importContent

https://github.com/apache/poi/blob/39a22a2c149a66a30c9fda0f266ffec934f3090c/src/ooxml/java/org/apache/poi/xslf/usermodel/XSLFShape.java#L116-L136

# Debugging Templating

The best way is to analyse the bfore and after.

```
  unpack the original template
  compare with the unpacked new (new1)
  use the new as a template (new1) to make new2
  compare new1 with new2
```

```
  # create the unpacked folders

  unzip -d orig original_file.pptx
  unzip -d new new_file.pptx

  mkdir old-new-compare
  find old-new-compare/
  find new -type f | xargs -L1 -Ixx sh -c 'cp xx old-new-compare/$(echo xx | sed "s/\//_/g")'
  find orig -type f | xargs -L1 -Ixx sh -c 'cp xx old-new-compare/$(echo xx | sed "s/\//_/g")'
  find old-new-compare -type f -name '*.xml' | xargs -L1 -Ixx sh -c 'mv xx xx.bak && xmllint --format xx.bak -o xx'
```

# Copying Slides

Some hints but no answers

http://apache-poi.1045710.n5.nabble.com/XSLF-copy-slides-td5147500.html
