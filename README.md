# funimage.figure
Make presentation slides and figures from Clojure for desktop software.

```
(-> (create-slideshow "~/tmp/slideshow001.pptx")
   (add-image (open-imp "~/Documents/camouflage_crab.tif"))
   (add-text "Crab" :x 200)
   (next-slide)
   (add-image (open-imp "~/Documents/camouflage_crab.tif") :x 100 :y 100)
   (add-text "Crab2" :x 200)
   (write-slideshow))
```
