##
#' Generic method to generate an APA style table with regression parameters for MS Word.
#'
#' @param  ... Regression (i.e., lm) result objects.
#' @param  variables The variable names for in the table.
#' @param  number (optional) The table number in the document.
#' @param  title (optional) Name of the table.
#' @param  filename (optional) Specify the filename (including valid '\code{.docx}' extension).
#' @param  note (optional) Add a footnote to the bottom of the table.
#' @param  landscape (optional) Set (\code{TRUE}) if the table should be generated in landscape mode.
#' @param  save (optional) Set (\code{FALSE}) if the table should not be saved in a document.
#' @param  type (optional) Not implemented.
#' @return \code{apa.regression} object; a list consisting of
#' \item{succes}{message in case of an error}
#' \item{save}{flag which indicates whther the document is saved}
#' \item{data}{dataset with regression statistics}
#' \item{table}{\code{FlexTable {ReporteRs}} object}
#' @importFrom "stats" "anova" "pf"
#' @export
##
apa.regression = function(..., variables=NULL, number="XX", title="APA Table", filename="APA Table.docx", note=NULL, landscape=FALSE, save=TRUE, type="wide") UseMethod("apa.regression")

##
#' Defualt method to generate an APA style table with regression parameters for MS Word.
#'
#' @param  ... Regression (i.e., lm) result objects.
#' @param  variables The variable names for in the table.
#' @param  number (optional) The table number in the document.
#' @param  title (optional) Name of the table.
#' @param  filename (optional) Specify the filename (including valid '\code{.docx}' extension).
#' @param  note (optional) Add a footnote to the bottom of the table.
#' @param  landscape (optional) Set (\code{TRUE}) if the table should be generated in landscape mode.
#' @param  save (optional) Set (\code{FALSE}) if the table should not be saved in a document.
#' @param  type (optional) Not implemented.
#' @return \code{apa.regression} object; a list consisting of
#' \item{succes}{message in case of an error}
#' \item{save}{flag which indicates whther the document is saved}
#' \item{data}{dataset with regression statistics}
#' \item{table}{\code{FlexTable {ReporteRs}} object}
#' @importFrom "stats" "anova" "pf"
#' @export
##
apa.regression.default = function(..., variables=NULL, number="XX", title="APA Table", filename="APA Table.docx", note=NULL, landscape=FALSE, save=TRUE, type="wide") {

  est = apaStyleRegression(list(...), variables, number, title, filename, note, landscape, save, type)
  est$call = match.call()
  class(est) = "apa.regression"
  est

}

##
#' Define a print method
#'
#' @param  x A \code{apa.regression} object
#' @export
##
print.apa.regression = function(x, ...) {
  if(x$succes == TRUE) {
    cat("\n")
    if (x$save == TRUE) {
      cat("Word document succesfully generated in: ")
      cat(getwd())
    } else {
      cat("Succesfully generated the APA table")
    }
    cat("\n\n")
  }
}

# The main function
apaStyleRegression = function(arguments, variables, number, title, filename, note, landscape, save, type) {

  # Initialize function
  options(warn = 0)

  regression.results = list()

  argument.type = lapply(arguments, class)

  # Only allow LM objects
  for (i in 1:length(arguments)) {
    if (argument.type[[i]] == "lm") {
      regression.results[[i]] = arguments[[i]]
    }
  }

  apa.report = c("B", "SE", "\u03B2", "*")
  intercept = FALSE
  steps = length(regression.results)

  # Internal function for calculating beta values (see also QuantPsych package)
  calculate.beta = function(model) {

    coef = summary(model)$coef[-1,1]
    sd.y = sapply(model$model[-1], sd)
    sd.x = sapply(model$model[1], sd)
    beta = c(coef * sd.x/sd.y)

    if (intercept == TRUE) {
      beta = c(NA, beta)
    }

    return(beta)
  }

  # Check the number of models specified by the user (Maximum of 3 are supported)
  if (steps > 3) {
    error = "A maximum of three hierachical models can be specified."
    warning(error)
    return(list(succes = error))
  }

  # Generate the variable names
  if(is.null(variables)) {

    # Create vector with variable names
    variable.names = names(regression.results[[steps]]$coefficients)
    variable.names = variable.names[-1]

    # Remove factor label
    if (length(regression.results[[steps]]$xlevels) > 0) {
      terms = paste(labels(regression.results[[steps]]$xlevels), "|", sep="", collapse="")
      terms = substr(terms, 1, nchar(terms)-1)
      variable.names = gsub(terms, "", variable.names, perl=TRUE)
    }

    variable.names = gsub("(\\w)(\\w*)", "\\U\\1\\L\\2", variable.names, perl=TRUE)

    if (intercept == TRUE) {
      variable.names = c("Intercept", gsub(":|\\."," x ", variable.names))
    }

  } else {

    # Check if length of supplied variables match that of the number of variables in the specified models
    nr.vars = ncol(regression.results[[steps]]$model) - 1
    if(nr.vars != length(variables)) {
      error = "The number of specified variables are different from the number of variables in the specified models."
      warning(error)
      return(list(succes = error))
    } else {
      variable.names = variables
    }

  }

  # Simple linear regression model
  if (steps == 1) {

    # Create parameters for a simple regression model
    summary.names = c("", "Model Summary", "\u0009df", "\u0009R\u00B2", "\u0009F")
    model.fit.vars = 5

    first.column = c(variable.names, summary.names)

    level2.header = level1.colspan = NULL
    level1.header = c("Variables", apa.report)

  } else {

    # Create parameters for a hierachical regression model #
    summary.names = c("", "Model Summary", "\u0009df", "\u0009R\u00B2", "\u0009F", "\u0009\u0394R\u00B2", "\u0009\u0394F")
    model.fit.vars = 7

    # Check all models have the same dependent variable
    same.dv = c()
    for (i in 1:steps) {
      same.dv[i] = colnames(regression.results[[1]]$model)[1] == colnames(regression.results[[i]]$model)[1]
    }
    if (any(same.dv==FALSE)) {
      error = "All specified models should have the same dependent variable."
      warning(error)
      return(list(succes = error))
    }


    # Check if the independent variables are specified in all models
    same.iv = TRUE
    for (i in 1:steps) {
      if (i < steps) {
        vars = ncol(regression.results[[i]]$model)
        for (j in 2:vars) {
          if (colnames(regression.results[[i]]$model)[j] != colnames(regression.results[[i+1]]$model)[j]) {
            same.iv = FALSE
          }
        }
      }
    }

    if (same.iv==FALSE) {
      error = "All specified regression models must contain the predictors from the preceeding models."
      warning(error)
      return(list(succes = error))
    }

    first.column = c(variable.names, summary.names)

    level1.header = c("", sapply(1:steps, function(x) paste("Model ", x)))
    level1.colspan = c(1, rep(4, steps))
    level2.header = c("Variables", rep(apa.report, steps))

  }

  statistics = list()
  r.change = c(rep(NA, steps))

  for (i in 1:steps) {

    # Extract model parameters
    model.summary = summary(regression.results[[i]])

    b = sprintf("%3.2f", round(model.summary$coefficients[2:nrow(model.summary$coefficients),1], digits = 2))
    b = c(b, rep("", model.fit.vars))
    b = c(b, rep("", (length(variable.names) - length(b)) + model.fit.vars))

    se = sprintf("%3.2f", round(model.summary$coefficients[2:nrow(model.summary$coefficients),2], digits = 2))
    se = c(se, rep("", model.fit.vars))
    se = c(se, rep("", (length(variable.names) - length(se)) + model.fit.vars))


    fit = c(sprintf("%3.2f", round(calculate.beta(regression.results[[i]]), digits = 2)))
    fit = c(fit, rep("", (length(variable.names) - length(fit)) + 2))

    sig = c(model.summary$coefficients[2:nrow(model.summary$coefficients),4])
    sig = c(sig, rep(NA, (length(variable.names) - length(sig)) + 2))

    f.stat = model.summary$fstatistic
    df = paste(f.stat[2], f.stat[3], sep=", ")
    r.value = model.summary$r.squared
    f.value = sprintf("%3.2f", round(f.stat[1], digits = 2))
    p.value = pf(f.stat[1], f.stat[2], f.stat[3], lower.tail=FALSE)

    fit = c(fit, df, sprintf("%3.2f", round(r.value, digits = 2)), f.value)
    sig = c(sig, NA, NA, p.value)

    # First model so change statistics are not available
    if (i == 1) {
      if (steps > 1) {
        fit = c(fit, rep("", 2))
        sig = c(sig, rep(NA, 2))
        r.change[i] = r.value
      }
    } else {
      r.change[i] = r.value - r.change[i-1]
      change = anova(regression.results[[i]], regression.results[[i-1]])
      fit = c(fit, sprintf("%3.2f", round(r.change[i], digits = 2)), sprintf("%3.2f", round(change$F[2], digits = 2)))
      sig = c(sig, NA, change[["Pr(>F)"]][2])
    }

    sig = ifelse(is.na(sig), "", ifelse(sig < .001, "***", ifelse(sig < .01, "**", ifelse(sig < .05, "*", ifelse(sig < .10, "\u2020", paste(c(rep("\u00A0", 6)), collapse = ""))))))

    statistics[[i]] = data.frame(b, se, fit, sig)
    colnames(statistics[[i]]) = c(paste("b", i, sep = "", collapse = ""), paste("se", i, sep = "", collapse = ""), paste("fit", i, sep = "", collapse = ""), paste("sig", i, sep = "", collapse = ""))

  }

  statistics = as.data.frame(statistics)
  statistics = cbind(variables = first.column, statistics)

  if (steps > 2) {
    landscape = TRUE
  }

  # Create Table
  apa.table = apaStyle::apa.table(data = data.frame(statistics), level1.header = level1.header, level1.colspan = level1.colspan, level2.header = level2.header, title = title, filename = filename, note = note, landscape = landscape, save = T)


  # Check if the document was succesfully generated
  if(apa.table$succes != TRUE) {
    return(list(succes = apa.table$succes))
  }

  return(list(succes = TRUE, save = save, data = data.frame(statistics), table = apa.table$table))

}
