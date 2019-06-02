### Web Scraping using Microsoft Excel VBA [Complete Course on Udemy]

#### Author : Kamal Girdher
#### Youtube : https://www.youtube.com/c/xtremeExcel
#### Telegram : https://t.me/letsautomate

--------------------------------------------------------------------------

## Index

 to be updated

--------------------------------------------------------------------------
## 1. Introduction & Disclaimer

### 1.1 What is Screen Scraping / Data Scraping?

Screen scraping is a practice of collecting all visual data from a website for use elsewhere. We typically automate the process by writing a script or using a software.

### 1.2 Is Scraping Legal?

It is not, unless you are scraping your own website or blog. We therefore do not promote Data Scraping in any way. Before you continue with this course, you accept that you won't misuse the content for any unauthorized or illegal activity.

--------------------------------------------------------------------------

## 2. VBA Refresher [Optional]

> **NOTE :** If you are not comfortable in VBA, refer [Excel macros/VBA Course on Youtube.](https://www.youtube.com/watch?v=dYHgr2murPk&list=PL1R_HJw0CDYKjmUxI3IKyuJcIKnWHVcuj)

### 2.1 Subprocedures & Functions

A **subprocedure** is a series of VB statements enclosed by the Sub and End Sub. It performs a task and then returns control to the calling code, but it does not return a value to the calling code.

```vba
Sub nameOfSubprocedure()
	<set of statments>
End Sub
```

A **function** on other hand is a series of VB statements enclosed by the Function and End Function. It performs a task and then returns control to the calling code. When it returns control, it also returns a value to the calling code.

```vba
Function nameOfFunction()
	<set of statments>
	nameOfFunction = <valueToBeReturned>
End Function
```

[Tutorial on Functions & Subprocedures](https://www.youtube.com/watch?v=1KDdu4BOZSA&list=PL1R_HJw0CDYKjmUxI3IKyuJcIKnWHVcuj&index=7&t=0s)


### 2.2 Variable Declaration and Object Initialization

**Variables** can store information required to use in our program. These are used to store values of various data types.

#### Declare a variable
```vba
Dim a As Integer

Dim b As String

Dim c As Variant
```

#### Initialize a variable
```vba
a = 10

b = "Kamal"
```

To understand variables in detail, refer these tutorials:
a. [Option Explicit & Implicit](https://www.youtube.com/watch?v=mojNkrnt_YA&list=PL1R_HJw0CDYKjmUxI3IKyuJcIKnWHVcuj&index=5&t=0s)
b. [Dim, Public, Private and Global Keywords](https://www.youtube.com/watch?v=33JmyY83IpA&list=PL1R_HJw0CDYKjmUxI3IKyuJcIKnWHVcuj&index=6&t=5s)


#### Declare an Object
```vba
Dim o as Object

Dim e As Excel.Application

Dim o As Outlook.Application
```

#### Initialize an Object
```vba
Set e = new Excel.Application

Set o = new Outlook.Application
```
--------------------------------------------------------------------------