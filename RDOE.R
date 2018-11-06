#' @title Design of Experiments Module
#'
#' @author Eduardo Castaño-Tostado, Nery Sofia Huerta-Pacheco, Victor Manuel Aguirre-Torres
#'
#' @description Module of Design of Experiments 'RDOE' is a user – oriented interface that integrates the functions from the book "Design and Analysis of Experiments with R" (Lawson, 2014 <ISBN: 9781439868133>) to fulfill the requirements of an experimental research considering a completely randomized one factor statistical design covering descriptive statistics, basic plots, anova, multiple comparisons tests and diagnostics. It will consider more complicate statistical designs in the future.
#'
#' You can learn more about this package at:
#' http://www.uv.mx/personal/nehuerta/rdoe/
#'
#' @details
#'
#' RDOE is a package with a graphical interface dedicated to perform of Design of Experiments in R.
#'
#' Note: RDOE is free software and comes with ABSOLUTELY NO WARRANTY.

#' @return RDOE is a graphic interface
#' @examples \dontrun{
#' ##Install package
#' library(RDOE)
#' ##Call the package
#' RDOE()
#' }
#'
#' @references
#' Lawson, J. (2014). Design and Analysis of Experiments with R.
#' CRC Press. ISBN: 9781439868133
#'
#' Lawson, J. (2016). daewr: Design and Analysis of Experiments with R.
#' R package version 1.1-7. https://CRAN.R-project.org/package=daewr
#'
#' Verzani, J. Based on the iwidgets code of Simon Urbanek, suggestions
#' by Simon Urbanek, Philippe Grosjean and Michael Lawrence (2014).
#' gWidgets: gWidgets API for building toolkit-independent, interactive GUIs.
#' R package version 0.0-54. https://CRAN.R-project.org/package=gWidgets
#'
#' @export RDOE
#' @import graphics
#' @import grDevices
#' @import utils
#' @import tcltk
#' @import gWidgets
#' @import readxl
#' @import REdaS
#' @import pander
#' @import multcomp
#' @import nortest
#' @import daewr
#' @importFrom stats TukeyHSD aov mad median model.tables sd var

RDOE<-function(){

  mi<- new.env()
  options("guiToolkit"="tcltk")

  ##Screen
  w<- gwindow("RDOE",visible=FALSE,width = 800,height= 510)
  g<- ggroup(horizontal=FALSE, spacing=0, container = w)

  nb <- gnotebook(container=g,width = 600,height= 500)
  g0<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "RDOE")
  g1<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Design")
  g3<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Descriptive")
  g4<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Inference Statistics")
  g5<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Checks")
  g6<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Comparison")
  g7<- ggroup(horizontal=FALSE, spacing=0, container = nb,label = "Power")

  #GLOBAL VARIABLES
  assign("gdata",NULL, envir =mi) #SI
  assign("gdata1",NULL, envir =mi)
  assign("Vari",NULL, envir =mi) #SI
  assign("a",NULL, envir =mi) #SI
  assign("b",NULL, envir =mi) #SI
  assign("varib",NULL, envir =mi) #SI
  assign("c1",NULL, envir =mi)#SI
  assign("c2",NULL, envir =mi)#SI
  assign("desing",NULL,envir=mi) #SI
  assign("desing1",NULL,envir=mi) #SI
  assign("desing2",NULL,envir=mi) #SI
  assign("na1",NULL, envir =mi) #SI
  assign("na2",NULL, envir =mi) #SI
  assign("Mat1",NULL, envir =mi) #SI
  assign("Mat2",NULL, envir =mi) #SI
  assign("Mat3",NULL, envir =mi) #SI
  assign("Mat4",NULL, envir =mi) #SI
  assign("plots",NULL, envir =mi) #SI
  assign("lp1",NULL, envir =mi) #SI
  assign("lp2",NULL, envir =mi) #SI
  assign("lp3",NULL, envir =mi) #SI
  assign("lp4",NULL, envir =mi) #SI
  assign("lp5",NULL, envir =mi) #SI
  assign("lp6",NULL, envir =mi) #SI

  ##MENU - OPEN
  #Open
  abrirc<-function(h,...){
    data<-tk_choose.files()
    data1<-read.csv(data)
    assign("gdata",data1, envir =mi)
  }

  abrirt<-function(h,...){
    data<-tk_choose.files()
    data1<-read.table(data,header=TRUE)
    assign("gdata",data1, envir =mi)
  }

  #Open xlsx
  openex<-function(h,...){
    data<-tk_choose.files()
    xlsx<-read_excel(data,sheet = 1, col_names=TRUE)
    data2<-as.data.frame(xlsx)
    assign("gdata",data2, envir =mi)
  }

  ##View
  ver<-function(h,...){
    gdata<-get("gdata",envir =mi)
    fix(gdata)
  }

  ##Re-start
  inicio<-function(h,...){
    dispose(w)
    RDOE()
  }

  #Save
  save1<-function(h,...){
    gdata1<-get("desing1",envir =mi)
    gdata2<-get("desing2",envir =mi)
    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(DOE)",".csv")
    nam1<-paste0(filename,"(DOE-Random)",".csv")
    write.csv(gdata1, file =nam)
    write.csv(gdata2, file =nam1)
  }


  #Save
  savedes<-function(h,...){
    Mat1x<-get("Mat1",envir =mi)
    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(Descriptive)",".csv")
    write.csv(Mat1x, file =nam)
  }

  #Save
  saveanv<-function(h,...){
    Mat2x<-get("Mat2",envir =mi)
    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(ANOVA)",".csv")
    write.csv(Mat2x, file =nam)
  }

  #Save
  savepower<-function(h,...){
    Mat3x<-get("Mat3",envir =mi)
    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(POWER)",".csv")
    write.csv(Mat3x, file =nam)
  }
  #Save
  savetable<-function(h,...){
    Mat4x<-get("Mat4",envir =mi)
    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(TUKEY.TABLE)",".csv")
    write.csv(Mat4x, file =nam)
  }



  #Save
  sadepl<-function(h,...){
    varib<-get("varib",envir =mi)
    c1<-get("c1",envir =mi)
    c2<-get("c2",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)
    a<-get("a",envir =mi)
    min<-min(a)
    max<-max(a)

    filename <- tclvalue(tkgetSaveFile())
    nam<-paste0(filename,"(Boxplot)",".png")

    png(filename=nam,width = 6, height = 6, units = 'in', res = 300)
    boxplot(varib[,c1]~varib[,c2],data=varib,col="lightgray",ylim=c(min-1,max+1),cex.axis=0.7,cex.lab=0.7,xlab=paste("Levels:",na2),ylab=paste("Response:",na1))
    dev.off()
  }

  #Save
  sachpl<-function(h,...){
    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    varib<-gdata
    n1<-na1
    n2<-na2
    n0<-ncol(gdata)

    for(v in 1:n0){
      if(n1==colnames(varib[v])){
        a<-gdata[,v]
        c1<-v
      }
    }
    assign("a",a, envir =mi)
    assign("c1",c1, envir =mi)

    for(v in 1:n0){
      if(n2==colnames(varib[v])){
        b<-gdata[,v]
        c2<-v
      }
    }
    assign("b",b, envir =mi)
    assign("c2",c2, envir =mi)

    min<-min(a)
    max<-max(a)
    Levels<-b
    Response<-a

    ##Case
    mod1 <- aov(Response ~ Levels)

    filename <- tclvalue(tkgetSaveFile())
    nam1<-paste0(filename,"(Residuals)",".png")
    nam2<-paste0(filename,"(Histogram)",".png")

    png(filename=nam1,width = 6, height = 6, units = 'in', res = 300)
    par(mfrow=c(2,2))
    plot(mod1,cex=.7,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1)
    dev.off()

    png(filename=nam2,width = 6, height = 6, units = 'in', res = 300)
    hist(mod1$residuals,main="",cex=.7,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1,prob=TRUE,xlab="Residuals")
    dev.off()

  }

  #Save
  satupl<-function(h,...){
    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    varib<-gdata
    n1<-na1
    n2<-na2
    n0<-ncol(gdata)

    for(v in 1:n0){
      if(n1==colnames(varib[v])){
        a<-gdata[,v]
        c1<-v
      }
    }
    assign("a",a, envir =mi)
    assign("c1",c1, envir =mi)

    for(v in 1:n0){
      if(n2==colnames(varib[v])){
        b<-gdata[,v]
        c2<-v
      }
    }
    assign("b",b, envir =mi)
    assign("c2",c2, envir =mi)

    min<-min(a)
    max<-max(a)
    Levels<-b
    Response<-a

    ##Case
    mod1 <- aov(Response ~ Levels)

    filename <- tclvalue(tkgetSaveFile())
    nam1<-paste0(filename,"(Tukey)",".png")

    png(filename=nam1,width = 6, height = 6, units = 'in', res = 300)
    plot(TukeyHSD(mod1,"Levels"),cex=.6,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1)
    dev.off()

  }

  ##Parameters

  #Desing
  parm0<-function(h,...){

    w1<- gwindow("Elements selection",visible=FALSE,width = 400,height= 350)
    tbl <- glayout(container=w1, horizontal=FALSE)
    var1<-c("Null",var)
    tbl[1,c(1:5)]<-(glx0<-glabel("Factor Design",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)

    tbl[2,c(3,4)]<-(glx2<-glabel("Factor",container=tbl))
    tbl[3,c(3,4)]<-(ge1<-gedit("",container=tbl))
    tbl[2,5]<-(gl1<-glabel("     Unit",container=tbl))
    tbl[3,5]<-(tg01<-gedit("",container=tbl))
    size(tg01)<-c(60,50)

    tbl[2,6]<-(gla2<-glabel("    Levels",container=tbl))
    tbl[2,7]<-(gla3<-glabel("Replication",container=tbl))
    tbl[3,6]<-(tg11<-gedit("",container=tbl))
    size(tg11)<-c(60,50)
    tbl[3,7]<-(tg21<-gedit("",container=tbl))
    size(tg21)<-c(60,50)

    tbl[10,5] <- (c <- gbutton("Save", container=tbl))
    addHandlerClicked(c, handler=function(h,...) {
      f1<-svalue(ge1)  #Factor name
      f2<-svalue(tg01) #Type
      f3<-svalue(tg11) #Levels
      f4<-svalue(tg21) #Rep
      le<-as.numeric(f3)-1
      len<-c(0:le)
      sam<-as.numeric(f3)*as.numeric(f4)
      tl<-paste0(f2,len)
      f<-rep(tl,each=as.numeric(f4))
      fac <- sample(f,sam) #Random
      ue<- 1:sam #exp unit
      mat0<-matrix(f,sam,1)
      colnames(mat0)<-c(f1)
      rownames(mat0)<-c(1:sam)

      mat1<-matrix(fac,sam,1)
      colnames(mat1)<-c(f1)
      rownames(mat1)<-c(1:sam)

      #
      m1<-as.data.frame(mat0)
      m2<-as.data.frame(mat1)
      desing1<-m1
      desing2<-m2
      assign("desing1",desing1, envir = mi)
      assign("desing2",desing2, envir = mi)
      save1()
    })

    tbl[10,6] <- (a <- gbutton("Cancel", container=tbl))
    addHandlerClicked(a, handler=function(h,...) {
      dispose(w1)
    })

    tbl[10,7] <- (b <- gbutton("Ok", container=tbl))

    addHandlerClicked(b, handler=function(h,...) {
      dispose(w1)

      f1<-svalue(ge1)  #Factor name
      f2<-svalue(tg01) #Type
      f3<-svalue(tg11) #Levels
      f4<-svalue(tg21) #Rep
      le<-as.numeric(f3)-1
      len<-c(0:le)
      sam<-as.numeric(f3)*as.numeric(f4)
      tl<-paste0(f2,len)
      f<-rep(tl,each=as.numeric(f4))
      fac <- sample(f,sam) #Random
      ue<- 1:sam #exp unit
      mat0<-matrix(f,sam,1)
      colnames(mat0)<-c(f1)
      rownames(mat0)<-c(1:sam)

      mat1<-matrix(fac,sam,1)
      colnames(mat1)<-c(f1)
      rownames(mat1)<-c(1:sam)

      #
      m1<-as.data.frame(mat0)
      m2<-as.data.frame(mat1)
      desing1<-m1
      desing2<-m2
      assign("desing1",desing1, envir = mi)
      assign("desing2",desing2, envir = mi)
      dispose(w1)
      DESING()
    })
    visible(w1) <- TRUE
  }

  #Input
  parmG<-function(h,...){
    gdata<-get("gdata",envir =mi)
    var<- colnames(gdata)
    n1<-ncol(gdata)
    w1<- gwindow("Elements selection",visible=FALSE,width = 400,height= 350)
    tbl <- glayout(container=w1, horizontal=FALSE)
    tbl[1,c(1:5)]<-(glx0<-glabel("Input",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[3,1] <- "Response"
    tbl[3,2] <- (cb1 <- gcombobox(var, container=tbl))
    tbl[3,3] <- "Factor"
    tbl[3,4] <- (cb2 <- gcombobox(var, container=tbl))
    size(cb1)<-c(50,50)
    size(cb2)<-c(50,50)
    tbl[5,3] <- (a <- gbutton("Cancel", container=tbl))
    addHandlerClicked(a, handler=function(h,...) {
      dispose(w1)
    })

    tbl[5,4] <- (b <- gbutton("Ok", container=tbl))

    addHandlerClicked(b, handler=function(h,...) {
      na1<-svalue(cb1)
      na2<-svalue(cb2)
      assign("na1",na1, envir =mi)
      assign("na2",na2, envir =mi)
      dispose(w1)
    })
    visible(w1) <- TRUE
  }

  #Descriptive
  parm1<-function(h,...){
    DESC()
  }


  #ANOVA
  parm2<-function(h,...){
    ANOVA()
  }

  #Check
  parm3<-function(h,...){
    CHECK()
  }

  parm4<-function(h,...){
    TUKEY()
  }

  parm5<-function(h,...){
    w1<- gwindow("Elements selection",visible=FALSE,width = 400,height= 350)
    tbl <- glayout(container=w1, horizontal=FALSE)
    var1<-c("Null",var)
    tbl[1,c(1:5)]<-(glx0<-glabel("Power",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)

    tbl[2,c(3,4)]<-(glx2<-glabel("Replications",container=tbl))
    tbl[3,3]<-(glx3<-glabel("Minimum",container=tbl))
    tbl[3,4]<-(ge1<-gedit("",container=tbl))
    size(ge1)<-c(60,50)
    tbl[4,3]<-(gl1<-glabel("Maximum",container=tbl))
    tbl[4,4]<-(tg01<-gedit("",container=tbl))
    size(tg01)<-c(60,50)

    tbl[1,6]<-(gla2<-glabel("Alpha",container=tbl))
    tbl[1,7]<-(tg11<-gedit("",container=tbl))
    size(tg11)<-c(60,50)
    tbl[2,6]<-(gla3<-glabel("Sigma",container=tbl))
    tbl[2,7]<-(tg12<-gedit("",container=tbl))
    size(tg12)<-c(60,50)
    tbl[3,6]<-(gla4<-glabel("Levels",container=tbl))
    tbl[3,7]<-(tg21<-gedit("",container=tbl))
    size(tg21)<-c(60,50)
    tbl[4,6]<-(gla5<-glabel("Delta",container=tbl))
    tbl[4,7]<-(tg22<-gedit("",container=tbl))
    size(tg22)<-c(60,50)


    tbl[10,6] <- (a <- gbutton("Cancel", container=tbl))
    addHandlerClicked(a, handler=function(h,...) {
      dispose(w1)
    })

    tbl[10,7] <- (b <- gbutton("Ok", container=tbl))

    addHandlerClicked(b, handler=function(h,...) {
      dispose(w1)

      f1<-svalue(ge1)  #Min
      f2<-svalue(tg01) #Max
      f3<-svalue(tg11) #Alpha
      f4<-svalue(tg12) #Sigma
      f5<-svalue(tg21)  #Levels
      f6<-svalue(tg22) #Delta
      l1<-as.numeric(f1)
      l2<-as.numeric(f2)
      l3<-as.numeric(f3)
      l4<-as.numeric(f4)
      l5<-as.numeric(f5)
      l6<-as.numeric(f6)

      assign("lp1",l1, envir =mi) #SI
      assign("lp2",l2, envir =mi) #SI
      assign("lp3",l3, envir =mi) #SI
      assign("lp4",l4, envir =mi) #SI
      assign("lp5",l5, envir =mi) #SI
      assign("lp6",l6, envir =mi) #SI

      dispose(w1)
      POWER()
    })
    visible(w1) <- TRUE
  }

  ##Cerrar
  cerrar<-function(h,...){
    dispose(w)
  }

  #Desing function
  DESI<-function(){
    matdes1<-get("desing1",envir = mi)
    matdes2<-get("desing2",envir = mi)

    m1<-as.data.frame(matdes1)
    m2<-as.data.frame(matdes2)
    desing<-list(m1,m2)
    names(desing) <- c("...DOE...","...DOE - Random...")
    assign("desing",desing,envir=mi)
    return(desing)
  }

  # 0 -------------------------------------------------------------------------------------------

  #Desing
  DESING<-function(h,...){
    matdes1<-get("desing1",envir = mi)
    matdes2<-get("desing2",envir = mi)

    DES<-function(){
      print(".....................................",quote=FALSE)
      print("Matrix desing",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      DESI()
    }

    tbl<-glayout(container=g1)
    gseparator(horizontal=TRUE, container=g1)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g1)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(DES())
    tbl[1,c(8:9)]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,c(8:9)]<-(b1<- gbutton("Tables",container=tbl,handler=save1))
    size(b1)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }

  # 1 -------------------------------------------------------------------------------------------
  #DESC
  DESC<-function(h,...){

    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    DES<-function(gdata,na1,na2){
      gdata<-get("gdata",envir =mi)
      na1<-get("na1",envir =mi)
      na2<-get("na2",envir =mi)

      varib<-gdata
      assign("varib",varib, envir =mi)
      n1<-na1
      n2<-na2
      n0<-ncol(gdata)

      for(v in 1:n0){
        if(n1==colnames(varib[v])){
          a<-gdata[,v]
          c1<-v
        }
      }
      assign("a",a, envir =mi)
      assign("c1",c1, envir =mi)

      for(v in 1:n0){
        if(n2==colnames(varib[v])){
          b<-gdata[,v]
          c2<-v
        }
      }
      assign("b",b, envir =mi)
      assign("c2",c2, envir =mi)

      min<-min(a)
      max<-max(a)

      ##Case
      d6<-t(tapply(varib[,c1], varib[,c2], mean))
      md6<-matrix(d6,1,)
      d7<-t(tapply(varib[,c1], varib[,c2], sd))
      md7<-matrix(d7,1,)
      d9<-t(tapply(varib[,c1], varib[,c2], median))
      md9<-matrix(d9,1,)
      d2<-t(tapply(varib[,c1], varib[,c2], mad))
      md2<-matrix(d2,1,)

      print(".....................................",quote=FALSE)
      print("Descriptive Statistics",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      print(paste("Response:",na1),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Levels:",na2),quote=FALSE)
      print("",quote=FALSE)
      print(".....................................",quote=FALSE)

      namef<-colnames(d6)
      mat<-round(rbind(md6,md7,md9,md2),digits=2)
      rownames(mat)<-c("Mean","SD","Median","Median absolute deviation")
      colnames(mat)<-c(namef)
      pandoc.table(mat, style = "grid",plain.ascii = TRUE)

      matdes<-as.data.frame(mat)
      Mat1<-matdes
      assign("Mat1",Mat1, envir = mi)

      #Graph
      dev.new()
      boxplot(varib[,c1]~varib[,c2],data=varib,col="lightgray",ylim=c(min-1,max+1),cex.axis=0.7,cex.lab=0.7,xlab=paste("Levels:",na2),ylab=paste("Response:",na1))
    }

    ##g3
    tbl<-glayout(container=g3)
    gseparator(horizontal=TRUE, container=g3)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g3)
    tbl <- glayout(container=gr1, horizontal=FALSE)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(DES(gdata,na1,na2))
    tbl[1,8]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,8]<-(b1<- gbutton("Table",container=tbl,handler=savedes))
    size(b1)<-c(70,50)
    tbl[3,8]<-(b2<- gbutton("Plot",container=tbl,handler=sadepl))
    size(b2)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }

  # 2 -------------------------------------------------------------------------------------------
  #ANOVA
  ANOVA<-function(h,...){

    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    ANOV<-function(gdata,na1,na2){
      gdata<-get("gdata",envir =mi)
      na1<-get("na1",envir =mi)
      na2<-get("na2",envir =mi)

      varib<-gdata
      n1<-na1
      n2<-na2
      n0<-ncol(gdata)

      for(v in 1:n0){
        if(n1==colnames(varib[v])){
          a<-gdata[,v]
          c1<-v
        }
      }
      assign("a",a, envir =mi)
      assign("c1",c1, envir =mi)

      for(v in 1:n0){
        if(n2==colnames(varib[v])){
          b<-gdata[,v]
          c2<-v
        }
      }
      assign("b",b, envir =mi)
      assign("c2",c2, envir =mi)

      min<-min(a)
      max<-max(a)
      Levels<-b
      Response<-a

      ##Case
      mod1 <- aov(Response ~ Levels)
      sum_test = unlist(summary(mod1))

      print(".....................................",quote=FALSE)
      print("Analysis Of Variance",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      print(paste("Response:",na1),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Levels:",na2),quote=FALSE)
      print("",quote=FALSE)
      print(".....................................",quote=FALSE)
      print(" ",quote=FALSE)
      a1<-sum_test["Df1"]
      a2<-sum_test["Df2"]
      a3<-sum_test["Sum Sq1"]
      a4<-sum_test["Sum Sq2"]
      a5<-sum_test["Mean Sq1"]
      a6<-sum_test["Mean Sq2"]
      a7<-sum_test["F value1"]
      a8<-sum_test["F value2"]
      a9<-sum_test["Pr(>F)1"]
      a10<-sum_test["Pr(>F)2"]
      a91<-formatC(a9, digits = 5, format = "f")

      a11<-a1+a2
      a22<-a3+a4
      a33<-a5+a6
      rmat<-round(c(a1,a2,a11,a3,a4,a22,a5,a6,a33,a7),digits=1)
      matano<-matrix(c(rmat,"-","-",a91,"-","-"),3,5)
      rownames(matano)<-c(na1,"Residuals","Total")
      colnames(matano)<-c("Df","Sum Sq","Mean Sq","F value","Pr(>F)")
      pandoc.table(matano,plain.ascii = TRUE)

      matanv<-as.data.frame(matano)
      Mat2<-matanv
      assign("Mat2",Mat2, envir = mi)

    }

    ##g4
    tbl<-glayout(container=g4)
    gseparator(horizontal=TRUE, container=g4)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g4)
    tbl <- glayout(container=gr1, horizontal=FALSE)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(ANOV(gdata,na1,na2))
    tbl[1,8]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,8]<-(b1<- gbutton("Table",container=tbl,handler=saveanv))
    size(b1)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }

  # 3 -------------------------------------------------------------------------------------------
  #CHECK
  CHECK<-function(h,...){

    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    CHE<-function(gdata,na1,na2){
      gdata<-get("gdata",envir =mi)
      na1<-get("na1",envir =mi)
      na2<-get("na2",envir =mi)

      varib<-gdata
      n1<-na1
      n2<-na2
      n0<-ncol(gdata)

      for(v in 1:n0){
        if(n1==colnames(varib[v])){
          a<-gdata[,v]
          c1<-v
        }
      }
      assign("a",a, envir =mi)
      assign("c1",c1, envir =mi)

      for(v in 1:n0){
        if(n2==colnames(varib[v])){
          b<-gdata[,v]
          c2<-v
        }
      }
      assign("b",b, envir =mi)
      assign("c2",c2, envir =mi)

      min<-min(a)
      max<-max(a)
      Levels<-b
      Response<-a

      ##Case
      mod1 <- aov(Response ~ Levels)

      print(".....................................",quote=FALSE)
      print("Anderson-Darling normality test",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      print(paste("Response:",na1),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Levels:",na2),quote=FALSE)
      print("",quote=FALSE)
      print(".....................................",quote=FALSE)
      x.test <- ad.test(mod1$residuals)  # Test de Lilliefors
      st<-round(x.test$statistic,digits=5)
      pv<-round(x.test$p.value,digits=5)
      cm<-c("A = ","p-value = ",st,pv)
      mcheck<-matrix(cm,2,)
      pandoc.table(mcheck,plain.ascii = TRUE)

      #Graph
      dev.new()
      par(mfrow=c(2,2))
      plot(mod1,cex=.7,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1)
      dev.new()
      hist(mod1$residuals,main="",cex=.7,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1,prob=TRUE,xlab="Residuals")

    }

    ##g3
    tbl<-glayout(container=g5)
    gseparator(horizontal=TRUE, container=g5)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g5)
    tbl <- glayout(container=gr1, horizontal=FALSE)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(CHE(gdata,na1,na2))
    tbl[1,8]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,8]<-(b1<- gbutton("Plots",container=tbl,handler=sachpl))
    size(b1)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }

  # 4 -------------------------------------------------------------------------------------------
  #TUKEY
  TUKEY<-function(h,...){

    gdata<-get("gdata",envir =mi)
    na1<-get("na1",envir =mi)
    na2<-get("na2",envir =mi)

    TUK<-function(gdata,na1,na2){
      gdata<-get("gdata",envir =mi)
      na1<-get("na1",envir =mi)
      na2<-get("na2",envir =mi)

      varib<-gdata
      n1<-na1
      n2<-na2
      n0<-ncol(gdata)

      for(v in 1:n0){
        if(n1==colnames(varib[v])){
          a<-gdata[,v]
          c1<-v
        }
      }
      assign("a",a, envir =mi)
      assign("c1",c1, envir =mi)

      for(v in 1:n0){
        if(n2==colnames(varib[v])){
          b<-gdata[,v]
          c2<-v
        }
      }
      assign("b",b, envir =mi)
      assign("c2",c2, envir =mi)

      min<-min(a)
      max<-max(a)
      Levels<-b
      Response<-a

      ##Case
      mod1 <- aov(Response ~ Levels)
      TuR<-TukeyHSD(mod1,"Levels") #Tukey en tablas
      A<-TuR$Levels
      A1<-formatC(A, digits = 5, format = "f")
      matdes<-as.data.frame(A1)
      Mat4<-matdes
      assign("Mat4",Mat4, envir = mi)

      TTab<-model.tables(mod1,"means")
      TSum<-summary(glht(mod1, linfct=mcp(Levels= "Dunnett"), alternative = "two.sided"))


      print(".....................................",quote=FALSE)
      print("Honestly-significant-difference",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      print(paste("Response:",na1),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Levels:",na2),quote=FALSE)
      print("",quote=FALSE)
      print("Note: The selection of Tukey's test or Dunnett's test should be based on the a priori comparisons of interest in the context of research; Tukey's test  should be used to compare all pairs of means while Dunnett's test should be used to compare means versus the one corresponding to the experimental control. The usage of both tests on the same data set is misleading.",quote=FALSE)
      print("",quote=FALSE)
      print(".....................................",quote=FALSE)
      print(" ",quote=FALSE)
      print("Tukey Test",quote=FALSE)
      print(" ",quote=FALSE)
      pandoc.table(A1, style = "grid",plain.ascii = TRUE)
      print(".....................................",quote=FALSE)
      print(" ",quote=FALSE)
      print("Estimated means",quote=FALSE)
      print(" ",quote=FALSE)
      print(TTab)
      print(".....................................",quote=FALSE)
      print(" ",quote=FALSE)
      print("Comparisons vs a control - Dunnett Contrasts",quote=FALSE)
      print(" ",quote=FALSE)
      print(TSum)
      #print(".....................................",quote=FALSE)
      #------------PLOT
      plot.TukeyHSD <- function (x, ...)
      {
        for (i in seq_along(x)) {
          xi <- x[[i]][, -4L, drop = FALSE] # drop p-values
          yvals <- nrow(xi):1L
          dev.hold(); on.exit(dev.flush())
          ## xlab, main are set below, so block them from ...
          plot(c(xi[, "lwr"], xi[, "upr"]), rep.int(yvals, 2L), type = "n",
               axes = FALSE, xlab = "", ylab = "", main = NULL, ...)
          axis(1, ...)
          axis(2, at = nrow(xi):1, labels = dimnames(xi)[[1L]], srt = 0, ...)
          abline(h = yvals, lty = 1, lwd = 0.5, col = "lightgray")
          abline(v = 0, lty = 2, lwd = 0.5, ...)
          segments(xi[, "lwr"], yvals, xi[, "upr"], yvals, ...)
          segments(as.vector(xi), rep.int(yvals - 0.1, 3L), as.vector(xi),
                   rep.int(yvals + 0.1, 3L), ...)
          title(main = paste0(format(100 * attr(x, "conf.level"), digits = 2L),
                              "% family-wise confidence level\n"),
                xlab = "Differences in mean levels")
          box()
          dev.flush(); on.exit()
        }
      }

      plot(TukeyHSD(mod1,"Levels"),cex=.6,cex.lab=.9,cex.axis=.9,cex.main=1, cex.sub=1) # Tukey gráfico
    }

    ##g4
    tbl<-glayout(container=g6)
    gseparator(horizontal=TRUE, container=g6)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g6)
    tbl <- glayout(container=gr1, horizontal=FALSE)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(TUK(gdata,na1,na2))
    tbl[1,8]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,8]<-(b1<- gbutton("Plots",container=tbl,handler=satupl))
    size(b1)<-c(70,50)
    tbl[3,8]<-(b2<- gbutton("Table",container=tbl,handler=savetable))
    size(b2)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }


  # 5 -------------------------------------------------------------------------------------------
  #POWER
  POWER<-function(h,...){

    lp1<-get("lp1",envir =mi)
    lp2<-get("lp2",envir =mi)
    lp3<-get("lp3",envir =mi)
    lp4<-get("lp4",envir =mi)
    lp5<-get("lp5",envir =mi)
    lp6<-get("lp6",envir =mi)

    POW<-function(lp1,lp2,lp3,lp4,lp5,lp6){
      lp1<-get("lp1",envir =mi)
      lp2<-get("lp2",envir =mi)
      lp3<-get("lp3",envir =mi)
      lp4<-get("lp4",envir =mi)
      lp5<-get("lp5",envir =mi)
      lp6<-get("lp6",envir =mi)

      rmin <- lp1
      rmax <- lp2

      alpha <- rep(lp3, rmax - rmin +1)
      sigma <- lp4
      nlev <- lp5
      Delta <- lp6
      nreps <- rmin:rmax
      power <- Fpower1(alpha, nlev, nreps, Delta, sigma)
      pa<-as.matrix(power)
      matdes<-as.data.frame(pa)
      Mat3<-matdes
      assign("Mat3",Mat3, envir = mi)

      print(".....................................",quote=FALSE)
      print("Power of a Statistical Test",quote=FALSE)
      print(".....................................",quote=FALSE)
      print("",quote=FALSE)
      print(paste("Minimum number of replicas:",lp1),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Maximum number of replicas:",lp2),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Significance level (Alpha):",lp3),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Estimation of standard deviation (Sigma):",lp4),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Levels:",lp5),quote=FALSE)
      print("",quote=FALSE)
      print(paste("Difference in mean (Delta):",lp6),quote=FALSE)
      print("",quote=FALSE)
      print(".....................................",quote=FALSE)
      print(" ",quote=FALSE)
      pandoc.table(pa,plain.ascii = TRUE)


    }

    ##g4
    tbl<-glayout(container=g7)
    gseparator(horizontal=TRUE, container=g7)
    gr1<- ggroup(horizontal=TRUE, spacing=0, container = g7)
    tbl <- glayout(container=gr1, horizontal=FALSE)
    tbl[c(1:10),c(1:7)]<-(outputArea <- gtext(container=tbl, expand=TRUE,width = 740,height= 350))
    out <- capture.output(POW(lp1,lp2,lp3,lp4,lp5,lp6))
    tbl[1,8]<-(glx0<-glabel("Save Elements",container=tbl))
    font(glx0) <- list(weight="bold",size= 9,family="sans",align ="center",spacing = 5)
    tbl[2,8]<-(b1<- gbutton("Table",container=tbl,handler=savepower))
    size(b1)<-c(70,50)
    dispose(outputArea)
    if(length(out)>0)
      add(outputArea, out)
  }



  #-------------------------------------------------------------------------------------------
  # MENUS
  abrir2<-list(one=gaction("csv",handler=abrirc),two=gaction("txt",handler=abrirt),three=gaction("xlsx",handler=openex))
  menulistaA<-list(Open=abrir2,u2=gaction("View",handler=ver),u3=gaction("Refresh",handler=inicio),u4=gaction("Close",handler=cerrar))
  imp<-list(cero=gaction("Design",handler=parm0),one=gaction("Input",handler=parmG),two=gaction("Descriptive",handler=parm1),three=gaction("ANOVA",handler=parm2),four=gaction("Checks",handler=parm3),five=gaction("Tukey-Dunnett",handler=parm4),six=gaction("Power",handler=parm5))

  #Manual
  y1<- function(h,..) gmessage("http://www.uv.mx/personal/nehuerta/rdoe/",title="Link")

  menulistaY<-list(u0=gaction("Manual",handler=y1))
  ##MENU
  mb_list <-list(File=menulistaA,CRD=imp,Help=menulistaY)
  gmenu(mb_list, container=g)

  ##g1
  #Information
  tmp1 <- gframe("", container=g0, expand=TRUE,horizontal=FALSE)
  tg<-glabel("                                                                  ",container=tmp1)
  tg<-glabel("               Design of Experiments in R              ",container=tmp1)
  font(tg) <- list(weight="bold",size= 26,family="sans",align ="center",spacing = 5)
  tg<-glabel("                                                                  ",container=tmp1)
  tg<-glabel("                                                                  ",container=tmp1)
  tg<-glabel("                                                                  ",container=tmp1)
  tg<-glabel("                                                                  ",container=tmp1)
  tg<-glabel("            UAQ                    UV                    ITAM",container=tmp1)
  font(tg) <- list(weight="bold",size= 24,family="sans",align ="center",spacing = 5)
  tg<-glabel("            Universidad Autonoma               Universidad Veracruzana               Statistics Department    ",container=tmp1)
  font(tg) <- list(weight="bold",size= 12,family="sans",align ="center",spacing = 5)
  tg<-glabel("                      de Queretaro",container=tmp1)
  font(tg) <- list(weight="bold",size= 12,family="sans",align ="center",spacing = 5)
  visible(w) <- TRUE
}
