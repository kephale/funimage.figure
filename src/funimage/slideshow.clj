(ns funimage.slideshow
  (:import [org.apache.poi.ss.usermodel Cell Row Sheet Workbook Chart ClientAnchor Drawing]
           [org.apache.poi.hssf.usermodel HSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [org.apache.poi.ss.usermodel.charts AxisCrosses AxisPosition ChartAxis ChartDataSource ChartLegend DataSources LegendPosition LineChartData ValueAxis]
           [org.apache.poi.ss.util CellRangeAddress]
           [java.io FileOutputStream])
  (:import [org.apache.poi.xslf.usermodel XMLSlideShow XSLFPictureData XSLFSlide XSLFPictureShape]
           [org.apache.poi.util IOUtils]
           [java.io FileInputStream FileOutputStream]
           [javax.imageio ImageIO]
           [java.awt Rectangle Color])     
  (:require [clojure.string :as string]
            [clojure.java.io :as io])
  (:use [funimage imp]))


(defn create-slideshow
   "Make a slideshow PPT and directory."
   [filename]
   (let [^XMLSlideShow ppt (XMLSlideShow. )
         ^XSLFSlide slide (.createSlide ppt)]
     ;(.mkdirs (java.io.File. directory))
     {:ppt ppt
      :slide slide
      :filename filename}))

(defn next-slide
  "Make a new slide."
  [slideshow]
  (assoc slideshow
         :slide (.createSlide (:ppt slideshow))))

(defn write-slideshow
  "Write the slideshow's ppt."
  [slideshow]
  (.write (:ppt slideshow)
    (FileOutputStream. (:filename slideshow)))
  slideshow)

; Add by writing to file first
#_(defn add-image
    "Add an image to the slideshow."
    [slideshow imp & args]
    (let [argmap (apply hash-map args)
          img-directory (str (:directory slideshow) java.io.File/separator "img")]
      (.mkdirs (java.io.File. img-directory))  
      (let [png-name (str img-directory java.io.File/separator (str (get-title imp) ".png"))           
            bi (ImageIO/write (.getBufferedImage imp)
                              "png"
                              (java.io.File. png-name))                     
            picture-data (IOUtils/toByteArray (FileInputStream. png-name))
            idx (.addPicture (:ppt slideshow) picture-data XSLFPictureData/PICTURE_TYPE_PNG)
            ^XSLFPictureShape picture-shape (.createPicture (:slide slideshow) idx)
            img-anchor (Rectangle. (or (:x argmap) 0) (or (:y argmap) 0)
                                   (get-width imp) (get-height imp))]
        (.setAnchor picture-shape img-anchor)))
    slideshow)

(defn add-image
    "Add an image to the slideshow."
    [slideshow imp & args]
    (let [argmap (apply hash-map args)]
      (let [;picture-data (.getPixels imp); might need to convert to byte processor first
            ; otherwise, convert to bufferedimage->imageio->bytearraywriter->bytes            
            baos (java.io.ByteArrayOutputStream.)
            wrote-bytes (javax.imageio.ImageIO/write (.getBufferedImage imp) "png" baos)
            picture-data (.toByteArray baos) 
            idx (.addPicture (:ppt slideshow) picture-data XSLFPictureData/PICTURE_TYPE_PNG)
            ^XSLFPictureShape picture-shape (.createPicture (:slide slideshow) idx)
            img-anchor (Rectangle. (or (:x argmap) 0) (or (:y argmap) 0)
                                   (or (:width argmap) (get-width imp)) (or (:height argmap) (get-height imp)))]
        (.setAnchor picture-shape img-anchor)))
    slideshow)
  
(defn add-text
  "Add text to the slideshow."
  [slideshow text & args]
  (let [argmap (apply hash-map args)
        txt-anchor (Rectangle. (or (:x argmap) 0) (or (:y argmap) 0)
                               (or (:width argmap) 100) (or (:height argmap) 100))
        text-box (.createTextBox (:slide slideshow))
        text-p (.addNewTextParagraph text-box)
        tr (.addNewTextRun text-p)]
    (.setText tr text)
    (.setFontColor tr (or (:font-color argmap) Color/black))
    (.setFontSize tr (or (:font-size argmap) 12))
    (.setAnchor text-box txt-anchor)
    slideshow))

#_(-> (create-slideshow "/Users/kyle/tmp/slideshow001.pptx")
   (add-image (open-imp "/Users/kyle/Documents/camouflage_crab.tif"))
   (add-text "Crab" :x 200)
   (next-slide)
   (add-image (open-imp "/Users/kyle/Documents/camouflage_crab.tif") :x 100 :y 100)
   (add-text "Crab2" :x 200)
   (write-slideshow))
  