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

## Issues

** fixed. New line bug - when the text being applied in has a newline, the style in the table, applied is wrong, except for the "last" entry (after the new line)
          Need to apply the styles to all textruns, in all paragraphs.
          adding new text (with new lines)  creates separate text runs
CONFLICT (modify/delete): docmaker/src/main/scala/inosion/DocMaker.scala deleted in HEAD and modified in origin/merge2.  Version origin/merge2 of docmaker/src/main/scala/inosion/DocMaker.scala left in tree.
CONFLICT (modify/delete): docmaker/src/main/scala/inosion/pptx/PPTXMerger.scala deleted in HEAD and modified in origin/merge2.  Version origin/merge2 of docmaker/src/main/scala/inosion/pptx/PPTXMerger.scala left in tree.
CONFLICT (modify/delete): docmaker/src/main/scala/inosion/pptx/Tools.scala deleted in HEAD and modified in origin/merge2.  Version origin/merge2 of docmaker/src/main/scala/inosion/pptx/Tools.scala left in tree.
CONFLICT (rename/rename): docmaker/test_pptx.py renamed to presenie/assets/test_pptx.py in HEAD and to presenie/samples/test_pptx.py in origin/merge2.
CONFLICT (file location): docmaker/assets/samplepptx.pptx renamed to docmaker/samples/sample_template1.pptx in origin/merge2, inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/sample_template1.pptx.
CONFLICT (file location): docmaker/samples/sample_template2.pptx added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/sample_template2.pptx.
CONFLICT (file location): docmaker/assets/test.zip added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test.zip.
CONFLICT (file location): docmaker/assets/test/TestXSLFBugs2.java added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test/TestXSLFBugs2.java.
CONFLICT (file location): docmaker/assets/test/closeview.jpeg added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test/closeview.jpeg.
CONFLICT (file location): docmaker/assets/test/dest.pptx added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test/dest.pptx.
CONFLICT (file location): docmaker/assets/test/longview.jpeg added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test/longview.jpeg.
CONFLICT (file location): docmaker/assets/test/source.pptx added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test/source.pptx.
CONFLICT (file location): docmaker/test_pptx.py renamed to docmaker/assets/test_pptx.py in origin/merge2, inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/samples/test_pptx.py.
CONFLICT (file location): docmaker/src/main/scala/org/apache/poi/xslf/usermodel/FilteredXSLFSheet.scala added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/src/main/scala/org/apache/poi/xslf/usermodel/FilteredXSLFSheet.scala.
Auto-merging presenie/src/main/scala/org/apache/poi/xslf/usermodel/RowBuilder.scala
CONFLICT (content): Merge conflict in presenie/src/main/scala/org/apache/poi/xslf/usermodel/RowBuilder.scala
CONFLICT (file location): docmaker/src/main/scala/org/apache/poi/xslf/usermodel/ShapeImporter.scala added in origin/merge2 inside a directory that was renamed in HEAD, suggesting it should perhaps be moved to presenie/src/main/scala/org/apache/poi/xslf/usermodel/ShapeImporter.scala.
