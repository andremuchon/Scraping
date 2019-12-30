# Scraping
Some small scraping projects

## 0. Code

All of the code is done in Python 3. Apart from what was used to move files around, the main libraries here are Pandas, Numpy, Selenium and Beautiful Soup.

## 1. Curva DI

### 1.1 Introduction

In Brazil, the CDI rate is an interest rate for overnight interbank deposits that aim to improve liquidity of financial institutions.

Every workday, B3 (Brazil Stock Exchange and Over-the-Counter Market) releases the market's position for the CDI rate at future dates, commonly called "vertices". This data is not available for every future workday, in a way that we need to interpolate between available vertices.

This interpolation of choice usually varies. Some manuals indicate that it exponential interpolation is preferred, while others argue in favor of cubic spline interpolation or even something called Nelson-Siegel-Svensson curve (Svensson, L. (1994). Estimating and Interpreting Foreward [sic] Interest Rates: Sweden 1992–1994, Papers 579 – Institute for International Economic Studies).

### 1.2 Project description
For this project, the goal is to scrap a brazilian interest curve (available online) and interpolate between the vertices available. The output format is an .xlsx file.

There are two versions of this, corresponding to different swaps of interest rates.

Since this is my first time working with some Python libraries, I chose to upload a "learn" notebook that recorded the process.

## 2. Debentures
Here, the goal is to periodically download and move files containing historical data about 'debentures' transactions, which available online. 

In order to run properly, the code needs to be modified to the appropriate output destination.
