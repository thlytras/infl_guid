ORtoRR <- function(OR, r) OR / (1-r + r*OR)
#HRtoRR <- function(HR, r) (1 - exp(HR*log(1-r)))/r
HRtoRR <- function(HR, r) sqrt((1 - (1-r)^HR)/(1 - (1-r)^(1/HR)))

SDp <- function(sd1, n1, sd2, n2) sqrt((sd1^2*(n1-1) + sd2^2*(n2-1))/(n1+n2-2))

# Load all the data in /input
library(readxl)
suppressMessages({
  dat.raw <- lapply(list.files("input", full=TRUE, pattern="\\.xlsx$"), function(x) {
    sapply(excel_sheets(x), function(y) as.data.frame(read_excel(x, y)), simplify=FALSE)
  })
})
names(dat.raw) <- gsub("(\\@)|( table)*\\.xlsx$", "", list.files("input", pattern="\\.xlsx$"))


# Prepping the data for analysis

# Seetting appropriate names
lblC <- c("author", "year", "pico", "design", "rob", "intervention", "comparator")
lblC2 <- c(lblC[1:5], "nos", lblC[6:7])
lblA <- paste0(c("n", "e", "pct"), rep(c(".i", ".c"), each=3))
lblB <- c("studlab", "n.c", "e.c", "n.i", "e.i")
lblB2 <- c("studlab", "intervention", "comparator", "n.c", "e.c", "n.i", "e.i")
lblT <- c("studlab", "intervention", "control", paste0(c("mean", "sd", "median", "q1", "q3"), rep(c(".i", ".c"), each=5)), "N.i", "N.c", "ci.lo.i", "ci.hi.i", "ci.lo.c", "ci.hi.c")
lblN1 <- c(paste0(c("N", "mean", "sd", "mean.lo", "mean.hi", "median", "q1", "q3", "median.lo", "median.hi"), rep(c(".i", ".c"), each=10)))
lblN2 <- c(paste0(c("N", "mean", "sd", "mean.lo", "mean.hi", "median", "q1", "q3", "iqr", "median.lo", "median.hi"), rep(c(".i", ".c"), each=11)))

dat1A <- dat.raw[["PICO 1A"]]
for (i in c(3, 5:6, 8:16)) {
  names(dat1A[[i]]) <- c(lblC, lblA)
}
names(dat1A[[1]]) <- c(lblC, lblN1, "aHR", "aHR.lo", "aHR.hi", "notes", "diff", "diff.hi", "diff.lo")
names(dat1A[[2]]) <- c(lblC, lblN2, "aHR", "aHR.lo", "aHR.hi")
names(dat1A[[4]]) <- c(lblC, lblA, "aOR", "aOR.lo", "aOR.hi", "aRR", "aRR.lo", "aRR.hi")
names(dat1A[[7]]) <- c(lblC, lblA, "aOR", "aOR.lo", "aOR.hi")


dat1B <- dat.raw[["PICO 1B"]]
names(dat1B[[3]]) <- c(lblC, lblA, "aOR", "aOR.lo", "aOR.hi", "aRR", "aRR.lo", "aRR.hi")
names(dat1B[[6]]) <- c(lblC, lblA, "aRR", "aRR.lo", "aRR.hi")
names(dat1B[[7]]) <- c(lblC, lblA, "aRR", "aRR.lo", "aRR.hi", "aOR", "aOR.lo", "aOR.hi")
for (i in c(4:5,8:16)) {
  names(dat1B[[i]]) <- c(lblC, lblA)
}
names(dat1B[[1]]) <- c(lblC, lblN1, "aHR", "aHR.lo", "aHR.hi", "notes")
names(dat1B[[2]]) <- c(lblC, lblN2, "aHR", "aHR.lo", "aHR.hi", "aOR", "aOR.lo", "aOR.hi")


dat2A <- dat.raw[["PICO 2A"]]
dat2A[length(dat2A)] <- NULL
dat2A[[3]][,14] <- NULL
names(dat2A[[4]]) <- c(lblC, lblA[-c(3,6)], "aOR", "aOR.lo", "aOR.hi")
names(dat2A[[7]]) <- c(lblC2, lblA, "aOR", "aOR.lo", "aOR.hi")
for (i in c(5:6,8:9)) {
  names(dat2A[[i]]) <- c(lblC2, lblA)
}
for (i in c(3,10:12)) {
  names(dat2A[[i]]) <- c(lblC, lblA)
}
names(dat2A[[1]]) <- c(lblC, "N.i", "N.c", "median.i", "median.c", "median.lo.i", "median.hi.i", "median.lo.c", "median.hi.c")
names(dat2A[[2]]) <- c(lblC2, "N.i", "N.c", "median.i", "median.c", "median.lo.i", "median.hi.i", "median.lo.c", "median.hi.c", "q1.i", "q3.i", "q1.c", "q3.c", "mean.i", "mean.c", "sd.i", "sd.c", "aHR", "aHR.lo", "aHR.hi", "aOR", "aOR.lo", "aOR.hi")



dat3A <- dat.raw[["PICO 3A"]]
for (i in c(1,3)) {
  names(dat3A[[i]]) <- c(lblB2[1:3], "inpatients", lblB2[4:7])
}
names(dat3A[[2]]) <- c(lblB2[1:3], "inpatients", lblB2[4:7], "aHR", "aHR.lo", "aHR.hi", "aOR", "aOR.lo", "aOR.hi", "Notes")
names(dat3A[[4]]) <- c(lblB, "aOR", "aOR.lo", "aOR.hi", "Notes")
for (i in c(5:12, 16, 19:22)) {
  names(dat3A[[i]]) <- lblB2
}
names(dat3A[[13]]) <- c(lblB2[c(1,4:5,2:3,6:7)], "aOR", "aOR.lo", "aOR.hi")
for (i in c(14:15)) {
  names(dat3A[[i]]) <- lblT
}
for (i in c(17:18)) {
  names(dat3A[[i]]) <- c(lblB2, "notes")
}
for (i in c(23:24)) {
  names(dat3A[[i]]) <- c(lblB, "notes")
}


dat3B <- dat.raw[["PICO 3B"]]
names(dat3B[[2]]) <- c(lblB, "aHR", "aHR.lo", "aHR.hi", "aOR", "aOR.lo", "aOR.hi", "Notes")
for (i in c(1, 3:7, 11, 13)) {
  names(dat3B[[i]]) <- lblB
}
for (i in c(8:10, 12)) {
  names(dat3B[[i]]) <- lblT[-(2:3)]
}



# Manual additions/adjustments

library(epitools)

i <- grep("Best", dat3A[[2]]$studlab) # Best 2024, death
dat3A[[2]][, c("aRR", "aRR.lo", "aRR.hi")] <- NA
dat3A[[2]][i, c("aRR", "aRR.lo", "aRR.hi")] <- riskratio(with(dat3A[[2]][i,], matrix(c(n.c-as.numeric(e.c), n.i-as.numeric(e.i), as.numeric(e.c), as.numeric(e.i)), nrow=2)))$measure[2,]
i <- grep("Best", dat3A[[4]]$studlab) # Best 2024, intubation
dat3A[[4]][, c("aRR", "aRR.lo", "aRR.hi")] <- NA
dat3A[[4]][i, c("aRR", "aRR.lo", "aRR.hi")] <- riskratio(with(dat3A[[4]][i,], matrix(c(n.c-as.numeric(e.c), n.i-as.numeric(e.i), as.numeric(e.c), as.numeric(e.i)), nrow=2)))$measure[2,]
i <- grep("Best", dat3A[[13]]$studlab) # Best 2024, hospitalization
dat3A[[13]][, c("aRR", "aRR.lo", "aRR.hi")] <- NA
dat3A[[13]][i, c("aRR", "aRR.lo", "aRR.hi")] <- riskratio(with(dat3A[[13]][i,], matrix(c(n.c-as.numeric(e.c), n.i-as.numeric(e.i), as.numeric(e.c), as.numeric(e.i)), nrow=2)))$measure[2,]




library(meta)
#library(estmeansd)

identifyObs <- function(d) {
  f <- lapply(d, function(x) names(x)[grep("^a..$", names(x))])
  names(f) <- paste(1:length(f), names(f))
  f[sapply(f, length)>0]
}

harmonize_effect_measures <- function(d) {
  n <- names(d)[grep("^a..$", names(d))]
  if (length(n)==0) return(d)
  if (length(n)==1) {
    d[,"est"] <- suppressWarnings(as.numeric(d[,n]))
    d[,"est.lo"] <- suppressWarnings(as.numeric(d[,paste0(n,".lo")]))
    d[,"est.hi"] <- suppressWarnings(as.numeric(d[,paste0(n,".hi")]))
    d$nest <- substr(n, 2,3)
    return(d)
  }
  r <- suppressWarnings(as.numeric(d$e.c) / as.numeric(d$n.c))
  d$est <- d$est.lo <- d$est.hi <- NA
  suppressWarnings({
  if ("aHR" %in% n) {
    d$est[!is.na(as.numeric(d$aHR))] <- HRtoRR(as.numeric(d$aHR[!is.na(as.numeric(d$aHR))]), r[!is.na(as.numeric(d$aHR))])
    d$est.lo[!is.na(as.numeric(d$aHR.lo))] <- HRtoRR(as.numeric(d$aHR.lo[!is.na(as.numeric(d$aHR.lo))]), r[!is.na(as.numeric(d$aHR.lo))])
    d$est.hi[!is.na(as.numeric(d$aHR.hi))] <- HRtoRR(as.numeric(d$aHR.hi[!is.na(as.numeric(d$aHR.hi))]), r[!is.na(as.numeric(d$aHR.hi))])
  }
  if ("aOR" %in% n) {
    d$est[!is.na(as.numeric(d$aOR))] <- ORtoRR(as.numeric(d$aOR[!is.na(as.numeric(d$aOR))]), r[!is.na(as.numeric(d$aOR))])
    d$est.lo[!is.na(as.numeric(d$aOR.lo))] <- ORtoRR(as.numeric(d$aOR.lo[!is.na(as.numeric(d$aOR.lo))]), r[!is.na(as.numeric(d$aOR.lo))])
    d$est.hi[!is.na(as.numeric(d$aOR.hi))] <- ORtoRR(as.numeric(d$aOR.hi[!is.na(as.numeric(d$aOR.hi))]), r[!is.na(as.numeric(d$aOR.hi))])
  }
  if ("aRR" %in% n) {
    d$est[!is.na(as.numeric(d$aRR))] <- as.numeric(d$aRR[!is.na(as.numeric(d$aRR))])
    d$est.lo[!is.na(as.numeric(d$aRR.lo))] <- as.numeric(d$aRR.lo[!is.na(as.numeric(d$aRR.lo))])
    d$est.hi[!is.na(as.numeric(d$aRR.hi))] <- as.numeric(d$aRR.hi[!is.na(as.numeric(d$aRR.hi))])
  }
  })
  d$nest <- "RR"
  return(d)
}


doMeta_bin <- function(d, t) {
  suppressWarnings({
    d$n.i <- as.numeric(d$n.i)
    d$e.i <- as.numeric(d$e.i)
    d$n.c <- as.numeric(d$n.c)
    d$e.c <- as.numeric(d$e.c)
  })
  d <- subset(d, !is.na(n.i) & n.i>0 & !is.na(n.c) & n.c>0 & n.i!="NA" & n.c!="NA")
  d <- subset(d, !((is.na(e.c) | e.c==0) & (is.na(e.i) | e.i==0))) # Trim zero event trials
  if (nrow(d)==0) return()
  if (!exists("studlab",d)) d$studlab <- with(d, paste0(author, ", ", year))
  if (!exists("year",d)) d$year <- as.numeric(substr(d$studlab, regexpr("([1-2][0-9][0-9][0-9])", d$studlab), regexpr("([1-2][0-9][0-9][0-9])", d$studlab)+3))
  if (exists("comparator",d)) d$comparator[d$comparator=="Untreated"] <- "Placebo"
  if (exists("design",d) && length(unique(d$design[!is.na(d$design)]))>1) {
    doMeta_bin(subset(d, design=="RCT"), paste(t, "RCTs"))
    doMeta_bin(subset(d, design=="non-RCT"), paste(t, "non-RCTs"))
    return()
  }
  if (exists("intervention")) {
    d$intervention[d$intervention=="Oseltamivir/peramivir"] <- "Oseltamivir/Peramivir" # Correct case
    d$intervention[d$intervention=="baloxavir"] <- "Baloxavir" # Correct case
  }
  cl <- paste(d$year, d$n.c, d$e.c)  # Identify clusters
  d$cl <- match(cl, unique(cl))
  cl2 <- paste(d$year, d$n.i, d$e.i)  # Identify clusters
  cl2 <- match(cl2, unique(cl2))
  if (max(cl2)<max(d$cl)) d$cl <- cl2
  has_sub <- function(x) {
    if (!exists(x,d)) return(FALSE)
    if (length(unique(d[,x][!is.na(d[,x])]))<=1) return(FALSE)
    #if (1 %in% sapply(unique(d[,x]), function(y) length(unique(d$cl[d[,x]==y])))) return(FALSE) # On second thought, this should not be applied...
    return(TRUE)
  }
  if (grepl("non(-){0,1}RCT", t)) {   # We have non-RCTs
    d <- harmonize_effect_measures(d)
    if (nrow(d)>1) {
      m <- try(metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=d, sm=d$nest[1], common=FALSE, cluster=cl, title=t), silent=TRUE)
      if (class(m)[1]=="try-error") m <- metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=d, sm=d$nest[1], common=FALSE, title=t)

      lcols <- c("studlab", "cluster"[!is.null(m$cluster)])
      cairo_pdf(paste0("output/", t, " - forest.pdf"), width=12, height=3+nrow(d)/3)
      forest(m, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, leftcols=lcols)
      dev.off()
      if (has_sub("rob")) {
        ms <- try(metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(studlab)), common=FALSE, cluster=cl, title=t, subgroup=rob), silent=TRUE)
        if (class(ms)[1]=="try-error") ms <- metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(studlab)), common=FALSE, title=t, subgroup=rob)
        lcols <- c("studlab", "cluster"[!is.null(ms$cluster)])
        cairo_pdf(paste0("output/", t, " - forest by ROB.pdf"), width=12, height=4+nrow(d)/3)
        forest(ms, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE, leftcols=lcols)
        dev.off()
      }
      if (has_sub("intervention")) {
        mi <- try(metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(intervention)), common=FALSE, cluster=cl, title=t, subgroup=intervention), silent=TRUE)
        if (class(mi)[1]=="try-error") mi <- metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(intervention)), common=FALSE, title=t, subgroup=intervention)
        lcols <- c("studlab", "cluster"[!is.null(mi$cluster)])
        cairo_pdf(paste0("output/", t, " - forest by treatment.pdf"), width=12, height=4+nrow(d)/3)
        forest(mi, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE, leftcols=lcols)
        dev.off()
      }
      if (has_sub("comparator")) {
        mc <- try(metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(comparator)), common=FALSE, cluster=cl, title=t, subgroup=comparator), silent=TRUE)
        if (class(mc)[1]=="try-error") mc <- metagen(log(est), (log(est.hi)-log(est.lo))/(-qnorm(0.025)*2), studlab, data=subset(d, !is.na(comparator)), common=FALSE, title=t, subgroup=comparator)
        lcols <- c("studlab", "cluster"[!is.null(mc$cluster)])
        cairo_pdf(paste0("output/", t, " - forest by comparator.pdf"), width=12, height=4+nrow(d)/3)
        forest(mc, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE, leftcols=lcols)
        dev.off()
      }
    }
  } else {  # We have RCTs
    if (nrow(d)>1) {
      m <- try(metabin(e.i, n.i, e.c, n.c, studlab, data=d, common=FALSE, cluster=cl, title=t), silent=TRUE)
      if (class(m)[1]=="try-error") m <- metabin(e.i, n.i, e.c, n.c, studlab, data=d, common=FALSE, title=t)

      cairo_pdf(paste0("output/", t, " - forest.pdf"), width=12, height=3+nrow(d)/3)
      forest(m, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE)
      dev.off()
      if (has_sub("rob")) {
        ms <- try(metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(rob)), common=FALSE, cluster=cl, title=t, subgroup=rob), silent=TRUE)
        if (class(ms)[1]=="try-error") ms <- metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(rob)), common=FALSE, title=t, subgroup=rob)
        cairo_pdf(paste0("output/", t, " - forest by ROB.pdf"), width=12, height=4+nrow(d)/3)
        forest(ms, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE)
        dev.off()
      }
      if (has_sub("intervention")) {
        mi <- try(metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(intervention)), common=FALSE, cluster=cl, title=t, subgroup=intervention), silent=TRUE)
        if (class(mi)[1]=="try-error") mi <- metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(intervention)), common=FALSE, title=t, subgroup=intervention)
        cairo_pdf(paste0("output/", t, " - forest by treatment.pdf"), width=12, height=4+nrow(d)/3)
        forest(mi, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE)
        dev.off()
      }
      if (has_sub("comparator")) {
        mc <- try(metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(comparator)), common=FALSE, cluster=cl, title=t, subgroup=comparator), silent=TRUE)
        if (class(mc)[1]=="try-error") mc <- metabin(e.i, n.i, e.c, n.c, studlab, data=subset(d, !is.na(comparator)), common=FALSE, title=t, subgroup=comparator)
        cairo_pdf(paste0("output/", t, " - forest by treatment.pdf"), width=12, height=4+nrow(d)/3)
        forest(mc, label.e="Intervention", label.c="Comparator", print.tau2=FALSE, print.pval.Q=FALSE, prediction=TRUE, prediction.subgroup=TRUE, test.subgroup=FALSE)
        dev.off()
      }
    }
  }
}  

cat("\nFitting PICO 1A: ")
for (i in c(3:16)) {
  cat(".")
  doMeta_bin(dat1A[[i]], paste("1A", names(dat1A)[i]))
}
cat("\nFitting PICO 1B: ")
for (i in c(3:16)) {
  cat(".")
  doMeta_bin(dat1B[[i]], paste("1B", names(dat1B)[i]))
}
cat("\nFitting PICO 2A: ")
for (i in c(3:12)) {
  cat(".")
  doMeta_bin(dat2A[[i]], paste("2A", names(dat2A)[i]))
}
cat("\nFitting PICO 3A: ")
for (i in c(1:13, 16:24)) {
  cat(".")
  doMeta_bin(dat3A[[i]], paste("3A", names(dat3A)[i]))
}
cat("\n\n")
cat("\nFitting PICO 3B: ")
for (i in c(1:7, 11, 13)) {
  cat(".")
  doMeta_bin(dat3B[[i]], paste("3B", names(dat3B)[i]))
}
cat("\n\n")



# Continuous outcomes:
# dat1A[c(1,2)]
# dat1B[c(1,2)]
# dat2A[c(1,2)]
# dat3A[16:17]
# dat3B[c(8:10, 12)]

# doMeta_cont <- function(d, t) {
#   d <- dat3A[[16]]
#   t <- names(dat3A)[16]
#
# }




# identifyObs(dat1A)
# identifyObs(dat1B)
# identifyObs(dat2A)
# identifyObs(dat3A)
# identifyObs(dat3B)



