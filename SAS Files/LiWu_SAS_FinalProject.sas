libname SAS "C:\00\00-DATA SCIENCE\15-SAS_Project\SAS_FinalProject\SAS_FinalProject_library"; 

************************* HEART DATASET ****************************************************************;
PROC IMPORT OUT= SAS.Heart 
            DATAFILE= "C:\00\00-DATA SCIENCE\15-SAS_Project\SAS_FinalPro
ject\heart.csv" 
            DBMS=CSV REPLACE;
     GETNAMES=YES;
     DATAROW=2; 
RUN;


/************************* FEATURES ********************************************************************;

Age: age of the patient [years]

Sex: sex of the patient [M: Male, F: Female]

ChestPainType: chest pain type [TA: Typical Angina, ATA: Atypical Angina, NAP: Non-Anginal Pain, ASY: Asymptomatic]

RestingBP: resting blood pressure [mm Hg]
		   Normal resting blood pressure for a healthy adult typically ranges from 90/60 mmHg to 120/80 mmHg.

Cholesterol: serum cholesterol [mm/dl]

FastingBS: fasting blood sugar [1: if FastingBS > 120 mg/dl, 0: otherwise]

RestingECG: resting electrocardiogram results 
           [Normal: Normal, 
            ST: having ST-T wave abnormality (T wave inversions and/or ST elevation or depression of > 0.05 mV), 
            LVH: showing probable or definite left ventricular hypertrophy by Estes' criteria]

MaxHR: maximum heart rate achieved [Numeric value between 60 and 202]

ExerciseAngina: exercise-induced angina [Y: Yes, N: No]

Oldpeak: oldpeak = ST [Numeric value measured in depression]
		[Thresholds for Concern
		<0 mm:May indicate heart attack or other conditions
		0 mm: No ST depression (normal).
		1-2 mm: Mild ST depression, may require further evaluation.
		>2 mm: Significant ST depression, often indicative of severe ischemia or coronary artery disease (CAD).]

ST_Slope: the slope of the peak exercise ST segment [Up: upsloping, Flat: flat, Down: downsloping]

TARGET:
HeartDisease: output class [1: heart disease, 0: Normal]
*/



/*Explore and describe the dataset briefly.*/
*Browsing Data Portion
Browsing Descriptor Portion;
%MACRO Brows(x);
TITLE "Browsing Data Portion of &x";
proc print data=SAS.Heart (obs=20);run;
TITLE "Browsing Descriptor Portion of &x";
proc contents data=SAS.Heart;run;
%MEND Brows;
%Brows(Heart Disease Dataset)




************************ NUMERICAL VARIABELS **************************************************************************
Age | RestingBP | Cholesterol | MaxHR | Oldpeak;

************************ CATEGORICAL VARIABELS ************************************************************************
Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope;



/*Duplicates in the dataset*/
proc sort data=SAS.Heart out=SAS.Heart_Sorted noduprecs dupout=SAS.Heart_Duplicates;
by _all_;
run;
proc contents data=SAS.Heart_Duplicates;
run;

/*Missing values in the dataset*/
*missing values in numerical variables;
TITLE "The Number of Missing Values in Numerical Variables";
proc means data=SAS.Heart nmiss;
run;

*Missing values in categorical variables;
TITLE "The Numeber of Missing Values in Categorical Variables";
proc freq data=SAS.Heart;
tables Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope / missing nocum;
run;



**************************************************************************************************
Univariate Analysis
**************************************************************************************************;
*NUMERICAL VARIABELS;
TITLE "Univariate Analysis: Numerical Variables in Heart Data";
proc univariate data=SAS.Heart normal plot;
var Age RestingBP Cholesterol MaxHR Oldpeak;
run;

/*CREATE MACRO FOR UNIVARITAE ANALYSIS (Numericable Variabels)*/
*Step 1;
TITLE "Univariate Analysis (Numerical): Age";
proc means data=SAS.Heart n nmiss min p10 q1 mean median q3 p90 p99 max std cv clm maxdec=2;
var Age;
run; 
proc sgplot data=SAS.Heart;
histogram Age / fillattrs=(color=lightblue);
density Age;
density Age / type=kernel;
run;
proc sgplot data=SAS.Heart;
hbox Age / fillattrs=(color=lightblue);
run;
ods select QQPlot TestsForNormality;
proc univariate data=SAS.Heart normal;
var Age;
qqplot Age / normal(mu=est sigma=est);
run;

*Step 2;
%let var=Age;
TITLE "Univariate Analysis (Numerical): &var";
proc means data=SAS.Heart n nmiss min p10 q1 mean median q3 p90 p99 max std cv clm maxdec=2;
var &var;
run; 
proc sgplot data=SAS.Heart;
histogram &var / fillattrs=(color=lightblue);
density &var;
density &var / type=kernel;
run;
proc sgplot data=SAS.Heart;
hbox &var / fillattrs=(color=lightblue);
run;
ods select QQPlot TestsForNormality;
proc univariate data=SAS.Heart normal;
var &var;
qqplot &var / normal(mu=est sigma=est);
run;

*Step 3;
%MACRO UniAnalysis_Num(var);
TITLE "Univariate Analysis (Numerical): &var";
/* Descriptive Statistics */
proc means data=SAS.Heart n nmiss min p10 q1 mean median q3 p90 p99 max std cv clm maxdec=2;
var &var;
run; 
/* Distribution */
proc sgplot data=SAS.Heart;
histogram &var / fillattrs=(color=lightblue);
density &var;
density &var / type=kernel;
run;
/* Boxplot: potential outliers */
proc sgplot data=SAS.Heart;
hbox &var / fillattrs=(color=lightblue);
run;
/* Normality Test */
ods select QQPlot TestsForNormality;
proc univariate data=SAS.Heart normal;
var &var;
qqplot &var / normal(mu=est sigma=est);
run;
%MEND UniAnalysis_Num;

**********  ALL NUMERICAL VARIABELS  **********;
%UniAnalysis_Num(Age);
%UniAnalysis_Num(RestingBP);
%UniAnalysis_Num(Cholesterol);
%UniAnalysis_Num(MaxHR);
%UniAnalysis_Num(Oldpeak);



****** OUTLIERS DETECTION (Turkey Menthod)******;
%MACRO Turkey_Outliers(var,threshold);
TITLE "&var Outlier Detection Based on Turkey Menthod Thresholds"; 
/*Step 1: Calculate Quartiles (Q1 and Q3)*/
proc means data=SAS.Heart noprint;
var &var;
output out=quartiles(drop=_TYPE_ _FREQ_) p25=Q1 p75=Q3;
run;
 /*Step 2: Compute IQR and Outlier Bounds*/
data outlier_bounds;
set quartiles;
IQR=Q3-Q1;
lower_bound=Q1-&threshold*IQR;
upper_bound=Q3+&threshold*IQR;
run;
/*Step 3: Detect Outliers*/
data detected_outliers;
if _n_=1 then set outlier_bounds;
set SAS.Heart;
if &var<lower_bound or &var>upper_bound then Outlier=1;
else Outlier=0;
run;
/*Step 4: Print Outliers*/
proc print data=detected_outliers;
where Outlier=1;
run;
%MEND Turkey_Outliers;

%Turkey_Outliers(Age,1.5);
%Turkey_Outliers(Age,3);

%Turkey_Outliers(RestingBP,1.5);  /*"0"*/
%Turkey_Outliers(RestingBP,3);

%Turkey_Outliers(Cholesterol,1.5); /*"0"*/
%Turkey_Outliers(Cholesterol,3);

%Turkey_Outliers(MaxHR,1.5);
%Turkey_Outliers(MaxHR,3);

%Turkey_Outliers(Oldpeak,1.5);
%Turkey_Outliers(Oldpeak,3);



****** OUTLIERS DETECTION (Z-Score)******;
%MACRO Std_Outliers(var,threshold);
TITLE "&var Outlier Detection Based on Z-Score Thresholds"; 
data Heart_outliers;
set SAS.Heart;
&var._original=&var;
run;
 /*Step 1: Standardize*/
proc stdize data=Heart_outliers out=Heart_Std;
var &var;
run;
/*Step 2: Detect Outliers*/
data &var._outliers;
set Heart_Std;
where &var < -&threshold or &var > &threshold;
run;
/*Step 3: Print Outliers*/
proc print data=&var._outliers;
where &var < -&threshold or &var > &threshold;
run;
%MEND Std_Outliers;

%Std_Outliers(Age,1.96)
%Std_Outliers(Age,3)

%Std_Outliers(RestingBP,1.96)
%Std_Outliers(RestingBP,3)

%Std_Outliers(Cholesterol,1.96)
%Std_Outliers(Cholesterol,3)

%Std_Outliers(MaxHR,1.96)
%Std_Outliers(MaxHR,3)

%Std_Outliers(Oldpeak,1.96)
%Std_Outliers(Oldpeak,3)



**********************************************************************************************
*CATEGORICAL VARIABELS;;
TITLE "Univariate Analysis: Categorical Variables in Heart Data";
proc freq data=SAS.Heart;
table Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope / nocum;
run;

/*CREATE MACRO FOR UNIVARITAE ANALYSIS (Categorical Variabels)*/
*Step 1;
TITLE "Univariate Analysis (Categorical): Sex";
proc freq data=SAS.Heart;
table Sex;
run;
goptions reset=all ftext="Arial" htext=1.2;
pattern1 color=lightcoral;
pattern2 color=lightblue;
pattern3 color=cream;
pattern4 color=lightgray;
proc gchart data=SAS.Heart;
pie sex / discrete
		  value=inside 
		  percent=inside
		  coutline=white;
run;
quit;

*Step 2;
%let var=Sex;
TITLE "Univariate Analysis (Categorical): &var";
proc freq data=SAS.Heart;
table &var;
run;
goptions reset=all ftext="Arial" htext=1.2;
pattern1 color=lightcoral;
pattern2 color=lightblue;
pattern3 color=cream;
pattern4 color=lightgray;
proc gchart data=SAS.Heart;
pie &var / discrete
		   value=inside 
		   percent=inside
		   coutline=white;
run;
quit;

*Step 3;
%MACRO UniAnalysis_Cat(var);
TITLE "Univariate Analysis (Categorical): &var";
/* Descriptive Statistics */
proc freq data=SAS.Heart;
table &var;
run;
/* Pie Chart */
goptions reset=all ftext="Arial" htext=1.2;
pattern1 color=lightcoral;
pattern2 color=lightblue;
pattern3 color=cream;
pattern4 color=lightgray;
proc gchart data=SAS.Heart;
pie &var / discrete
		   value=inside 
		   percent=inside
		   coutline=white;
run;
quit;
%MEND UniAnalysis_Cat;

**********  ALL CATEGORICAL VARIABELS  **********;
%UniAnalysis_Cat(Sex);
%UniAnalysis_Cat(ChestPainType);
%UniAnalysis_Cat(FastingBS);
%UniAnalysis_Cat(RestingECG);
%UniAnalysis_Cat(ExerciseAngina);
%UniAnalysis_Cat(ST_Slope);


/*Target Variable: HeartDisease*/
%MACRO UniAnalysis_Cat_Vbar(var);
TITLE "&var";
/* Descriptive Statistics */
proc freq data=SAS.Heart;
table &var / out=FreqData;
run;
ods graphics / width=800px height=600px;
/* Bar Chart */
proc sgplot data=FreqData;
    vbar &var / response=COUNT datalabel 
                datalabelattrs=(size=12) 
                barwidth=0.6
                fillattrs=(color=lightcoral)
                outlineattrs=(color=black);
    yaxis label="Frequency" grid;
    xaxis display=(nolabel);
run;
quit;
%MEND UniAnalysis_Cat_Vbar;

%UniAnalysis_Cat_Vbar(HeartDisease);


***********************************************************************************************************
Bivariate Analysis
***********************************************************************************************************
Continouse Vs. Continouse  : For Visulaization scatter plot,...
                             For test of independence: pearson correlation or spearman or Kendal tau, ...
            
Categorical Vs. Categorical: For summaraization: contingency table (two-way table)
                             For visualization :stacked bar chart,Grouped bar chart,...
                             For test of independence:chi-square test

Continouse Vs. Categorical : For summaraization:gropup by categorical column an aggragte for numerical column
                             For visualization: Grouped box plot,Grouped histogram,Grouped density,...
                             For test of independence :1) if categorical column has only two levels :t-test
                                                          2) if categorical column has more than two levels: ANOVA
*/



*******************************************************************************************************
Categorical vs. Categorical (Chi-Square Test);
/*Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope | VS  HeartDisease*/ 

%MACRO Cat_Cat(var1, var2);
TITLE "Stacked Grouped Bar Chart of &var1 by &var2";
proc sgplot data=SAS.Heart;
styleattrs datacolors=(lightblue lightcoral);
vbar &var1 / group=&var2 datalabel;
run;
TITLE "Chi-Square Test for &var1 and &var2";
proc freq data=SAS.Heart;
table &var1*&var2 / chisq;
run;
%MEND Cat_Cat;


/*Sex vs HeartDisease*/
%Cat_Cat(Sex, HeartDisease);

proc logistic data=SAS.Heart;
class Sex (ref="F")/param=glm;
model HeartDisease(event="1")=Sex;
lsmeans sex / e ilink;
run; 

/*ChestPainType vs HeartDisease*/
%Cat_Cat(ChestPainType, HeartDisease);

proc logistic data=SAS.Heart;
class ChestPainType (ref="ATA")/param=glm;
model HeartDisease(event="1")=ChestPainType;
lsmeans ChestPainType / e ilink;
run;

/*FastingBS vs HeartDisease*/
%Cat_Cat(FastingBS, HeartDisease);

proc logistic data=SAS.Heart;
class FastingBS (ref="0")/param=glm;
model HeartDisease(event="1")=FastingBS;
lsmeans FastingBS / e ilink;
run;

/*RestingECG vs HeartDisease*/
%Cat_Cat(RestingECG, HeartDisease);

proc logistic data=SAS.Heart;
class RestingECG (ref="Normal")/param=glm;
model HeartDisease(event="1")=RestingECG;
lsmeans RestingECG / e ilink;
run;

/*ExerciseAngina vs HeartDisease*/
%Cat_Cat(ExerciseAngina, HeartDisease);

proc logistic data=SAS.Heart;
class ExerciseAngina (ref="N")/param=glm;
model HeartDisease(event="1")=ExerciseAngina;
lsmeans ExerciseAngina / e ilink;
run;

/*ST_Slope vs HeartDisease*/
%Cat_Cat(ST_Slope, HeartDisease);

proc logistic data=SAS.Heart;
class ST_Slope (ref="Up")/param=glm;
model HeartDisease(event="1")=ST_Slope;
lsmeans ST_Slope / e ilink;
run;



*************************************************************************************************************
Continouse vs. Categorical (Target "HeartDisease" is Categorical Variable);
/*Age | RestingBP | Cholesterol | MaxHR | Oldpeak | VS HeartDisease*/

%MACRO Con_Cat(var);
title "Impact of &var on Heart Disease";
proc logistic data=SAS.Heart;
model HeartDisease(event='1')=&var;
run;
/*Visuliation:Box Plot*/
proc sgplot data=SAS.Heart;
vbox &var / category=HeartDisease
			fillattrs=(color=lightblue);
run;
%MEND Con_Cat;

%Con_Cat(Age);
%Con_Cat(RestingBP);
%Con_Cat(Cholesterol);
%Con_Cat(MaxHR);
%Con_Cat(Oldpeak);



*******************************************************************************************************
AMONG ALL NUMERICAL FEATURES;
*Continouse Vs. Continouse;
/*Age | RestingBP | Cholesterol | MaxHR | Oldpeak*/

TITLE "Conducting Correaltion Between Numerical Variables";
proc corr data=SAS.Heart_imp pearson spearman plots(maxpoints=none)=matrix(histogram);
var Age RestingBP Cholesterol MaxHR Oldpeak;
run;



*******************************************************************************************************
AMONG ALL CATEGORIAL FEATURES;
*Categorical vs. Categorical (Chi-Square Test);
/*Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope*/

proc freq data=SAS.Heart;
table Sex*(ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope) /chisq;
run;

proc freq data=SAS.Heart;
table ChestPainType*(FastingBS RestingECG ExerciseAngina ST_Slope) /chisq;
run;

proc freq data=SAS.Heart;
table FastingBS*(RestingECG ExerciseAngina ST_Slope) /chisq;
run;

proc freq data=SAS.Heart;
table RestingECG*(ExerciseAngina ST_Slope) /chisq;
run;

proc freq data=SAS.Heart;
table ExerciseAngina*ST_Slope /chisq;
run;



*******************************************************************************************************
AMONG NUMERICAL & CATEGORICAL FEATURES;
/*Continouse vs. Categorical (Target is Continuous Variable)*/
*Con: Age | RestingBP | Cholesterol | MaxHR | Oldpeak |
*Cat: Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope |;
TITLE;
%MACRO Cat_Con(var,target);
proc glm data=SAS.Heart;
class &var;
model &target=&var;
means &var / hovtest=levene(type=abs) welch;
run;
%MEND Cat_Con;

%MACRO Features_Con_Cat(con,cat);
proc logistic data=SAS.Heart;
class &cat (param=ref);  			/* Reference coding for categorical variable */
model &cat = &con / link=glogit; 
run;
%MEND Features_Con_Cat;


/*Sex(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(Sex,RestingBP);
%Cat_Con(Sex,Cholesterol);
%Cat_Con(Sex,MaxHR);
%Cat_Con(Sex,Oldpeak);


/*ChestPainType(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(ChestPainType,RestingBP);
%Cat_Con(ChestPainType,Cholesterol);
%Cat_Con(ChestPainType,MaxHR);
%Cat_Con(ChestPainType,Oldpeak);

%Features_Con_Cat(Cholesterol,ChestPainType);
%Features_Con_Cat(MaxHR,ChestPainType);
%Features_Con_Cat(Oldpeak,ChestPainType);
%Features_Con_Cat(Age,ChestPainType);



/*FastingBS(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(FastingBS,RestingBP);
%Cat_Con(FastingBS,Cholesterol);
%Cat_Con(FastingBS,MaxHR);
%Cat_Con(FastingBS,Oldpeak);

%Features_Con_Cat(Age,FastingBS);


/*RestingECG(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(RestingECG,RestingBP);
%Cat_Con(RestingECG,Cholesterol);
%Cat_Con(RestingECG,MaxHR);
%Cat_Con(RestingECG,Oldpeak);

%Features_Con_Cat(RestingBP,RestingECG);
%Features_Con_Cat(Cholesterol,RestingECG);
%Features_Con_Cat(Age,RestingECG);



/*ExerciseAngina(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(ExerciseAngina,RestingBP);
%Cat_Con(ExerciseAngina,Cholesterol);
%Cat_Con(ExerciseAngina,MaxHR);
%Cat_Con(ExerciseAngina,Oldpeak);

%MACRO Angina_Con_Cat(var);
proc logistic data=SAS.Heart;
model ExerciseAngina(event='Y')=&var;
run;
/*Visuliation:Box Plot*/
proc sgplot data=SAS.Heart;
vbox &var / category=ExerciseAngina
			fillattrs=(color=lightblue);
run;
%MEND Angina_Con_Cat;

%Angina_Con_Cat(RestingBP);
%Angina_Con_Cat(MaxHR);
%Angina_Con_Cat(Oldpeak);
%Angina_Con_Cat(Age);



/*ST_Slope(Categorical) vs. Continuous*/
*RestingBP | Cholesterol | MaxHR | Oldpeak;
%Cat_Con(ST_Slope,RestingBP);
%Cat_Con(ST_Slope,Cholesterol);
%Cat_Con(ST_Slope,MaxHR);
%Cat_Con(ST_Slope,Oldpeak);

%Features_Con_Cat(RestingBP,ST_Slope);
%Features_Con_Cat(Cholesterol,ST_Slope);
%Features_Con_Cat(MaxHR,ST_Slope);
%Features_Con_Cat(Oldpeak,ST_Slope);
%Features_Con_Cat(Age,ST_Slope);



********************************************** General Analysis **********************************************;
ods graphics on;
proc pls data=SAS.Heart plots=all;
class HeartDisease Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope;
model Cholesterol = HeartDisease Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope 
					Age RestingBP MaxHR Oldpeak / solution;
run;
quit;
ods graphics off;

*Check MultiCollinearity Among Numerical Variables;
ods graphics on;
proc reg data=SAS.Heart;
model HeartDisease=Age RestingBP Cholesterol MaxHR Oldpeak / vif collinoint;
output out = outstat 
		p = Predicted 
		r = Residual 
		stdr = se_resid 
		rstudent = RStudent 
		h = Leverage 
		cookd = CooksD; 
run;
quit;
ods graphics off;



*********************************************************************************************************************
Explaination using Logistic Regression;
ods graphics on; 
proc logistic data=SAS.Heart plots(only)=(effect oddsratio); 
class Sex(ref="F") ChestPainType(ref="ATA") FastingBS(ref="0") RestingECG(ref="ST") 
	  ExerciseAngina(ref="N") ST_Slope(ref="Up")/param=ref;
model HeartDisease(event="1")= Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope 
	  Age RestingBP Cholesterol MaxHR Oldpeak / details lackfit; 
output out=pred p=phat lower=lcl upper=ucl predprob=(individual crossvalidate);
ods output Association=Association; 
run; 
quit;
ods graphics off;



*************************************************************************************************************
/*PREDICTIVE MODELING*/
*************************************************************************************************************
Handling Missing Values (RestingBP, Cholesterol)


Number of RestingBP=0;
proc sql;
    select count(*) as num_obs
    from SAS.Heart
    where RestingBP=0;
quit;
*Impute RestingBP missing values with Median(130.00);
data SAS.Heart_imp;
set SAS.Heart;
if RestingBP=0 then RestingBP=130;
run;


*Number of Cholesterol=0;
proc sql;
    select count(*) as num_obs
    from SAS.Heart
    where Cholesterol=0;
quit;

*Impute Cholesterol missing values using mice imputation;
title;
data SAS.Heart_imp;
set SAS.Heart_imp;
if Cholesterol=0 then Cholesterol=.;
imp_Cholesterol_mean=Cholesterol;
imp_Cholesterol_median=Cholesterol;
imp_Cholesterol_mice=Cholesterol;
 if Cholesterol=. then imp_Cholesterol_mean=198.80;
 if Cholesterol=. then imp_Cholesterol_median=223.00;
run;
proc print data=SAS.Heart_imp;
run;

*Check correlation;
proc corr data=SAS.Heart_imp pearson spearman;
var Age RestingBP MaxHR Oldpeak Cholesterol imp_Cholesterol_mean imp_Cholesterol_median;
run;

ods select misspattern;
proc mi data=SAS.Heart_imp nimpute=0;
class Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope;
fcs logistic regpmm;
var Age Sex ChestPainType RestingBP FastingBS RestingECG MaxHR ExerciseAngina Oldpeak ST_Slope imp_Cholesterol_mice;
run;

*Impute;
proc mi data=SAS.Heart_imp nimpute=1 seed=2025 out=SAS.Heart_imp;
class Sex ChestPainType FastingBS RestingECG ExerciseAngina ST_Slope;
fcs logistic regpmm;
var Age Sex ChestPainType RestingBP FastingBS RestingECG MaxHR ExerciseAngina Oldpeak ST_Slope imp_Cholesterol_mice;
run;
*Check correlation Again;
proc corr data=SAS.Heart_imp pearson spearman;
var Age RestingBP MaxHR Oldpeak Cholesterol imp_Cholesterol_mice imp_Cholesterol_mean imp_Cholesterol_median;
run;
proc print data=SAS.Heart_imp;
run;


*Imputed Cholesterol association with Target Variable (HeartDisease);
%MACRO Con_Cat_imp(var);
proc logistic data=SAS.Heart_imp;
model HeartDisease(event='1')=&var;
run;
%MEND Con_Cat_imp;
%Con_Cat_imp(imp_Cholesterol_mice);


*Binning ChestPainType (NAP+TA="NAP+TA");
*Binning ST_Slope (Down+Flat="Down+Flat");
data SAS.Heart_imp;
set SAS.Heart_imp;
length ChestPainType_N $10;
length ST_Slope_N $10;
ChestPainType_N = ChestPainType;
ST_Slope_N = ST_Slope;
if ChestPainType in ("NAP", "TA") then ChestPainType_N="NAP+TA";
if ST_Slope in ("Down", "Flat") then ST_Slope_N="Down+Flat";
run;
proc freq data=SAS.Heart_imp;
table ChestPainType_N ST_Slope_N;
run;
proc print data=SAS.Heart_imp;
where ST_Slope="Down" or ST_Slope="Flat";
run;


*********************************************************************************************************************
After handling missing values and binning of ChestPainType
Explaination using Logistic Regression;
ods graphics on; 
proc logistic data=SAS.Heart_imp plots(only)=(effect oddsratio); 
class Sex(ref="F") ChestPainType_N(ref="ATA") FastingBS(ref="0") RestingECG(ref="ST") 
	  ExerciseAngina(ref="N") ST_Slope_N(ref="Up")/param=ref;
model HeartDisease(event="1")= Sex ChestPainType_N FastingBS RestingECG ExerciseAngina ST_Slope_N 
	  Age RestingBP Cholesterol MaxHR Oldpeak / details lackfit; 
output out=pred p=phat lower=lcl upper=ucl predprob=(individual crossvalidate);
ods output Association=Association; 
run; 
quit;
ods graphics off;



*************************************************************************************************************
LOGISTIC REGRESSION (HeartDisease)
*************************************************************************************************************
*Con: Age | RestingBP | Cholesterol | MaxHR | Oldpeak |
*Cat: Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope |;


*Split Dataset into Train set & Test set;
proc surveyselect data=SAS.Heart_imp rate=0.90 outall out=result seed=2025; 
run;
*check distribution of sampling;
proc freq data=result;
table Selected;
run;
data traindata testdata;
set result;
if selected=1 then output traindata;
else output testdata;
run;



*Logistic Regression;
ods html;
ods graphics on; 
proc logistic data=traindata plots=(ROC); 
    class Sex(ref="F") ChestPainType_N(ref="ATA") FastingBS(ref="0") RestingECG(ref="ST") ExerciseAngina(ref="N") 
		  ST_Slope_N(ref="Up") / param=ref;    
    model HeartDisease(event="1") = Sex ChestPainType_N FastingBS RestingECG ExerciseAngina ST_Slope_N 
                                    Age RestingBP imp_Cholesterol_mice MaxHR Oldpeak / details lackfit;  
    /* Score the training dataset */
    score data=traindata out=trainpred;  
    /* Score the test dataset */
    score data=testdata out=testpred outroc=roc_logistic;  
    roc; 
    roccontrast; 
    /* Output probabilities */
    output out=outputdata p=prob_predicted xbeta=linpred;
run; 
quit;
ods graphics off;


* Confusion matrix, Recall(Sensitivity), Specificity;
/*Train Set*/
ods html;
ods html on;
ods html style=journal;
proc sort data=trainpred;
by descending I_HeartDisease descending F_HeartDisease;
run;
proc freq data=trainpred order=data;
tables F_HeartDisease*I_HeartDisease / senspec;
run;
ods html close;

/*Test Set*/
ods html on;
ods html style=journal;
proc sort data=testpred;
by descending I_HeartDisease descending F_HeartDisease;
run;
proc freq data=testpred order=data;
tables F_HeartDisease*I_HeartDisease / senspec;
run;
ods html close;



*************************************************************************************************************
Decision Tree (HeartDisease)
*************************************************************************************************************
*Con: Age | RestingBP | Cholesterol | MaxHR | Oldpeak |
*Cat: Sex | ChestPainType | FastingBS | RestingECG | ExerciseAngina | ST_Slope |;

*Decision Tree;
ods html;
ods graphics on; 
proc HPSPLIT data=traindata; 
	class HeartDisease Sex ChestPainType_N FastingBS RestingECG ExerciseAngina ST_Slope_N;
	model HeartDisease(event="1")=Sex ChestPainType_N FastingBS RestingECG ExerciseAngina ST_Slope_N 
		  Age RestingBP imp_Cholesterol_mice MaxHR Oldpeak; 
	prune costcomplexity;
	code file="C:\00\00-DATA SCIENCE\15-SAS_Project\SAS_FinalProject\treeoutput.sas";
	/* Save Predictions */
    output out=scored;
run; 
quit;
ods graphics off;
ods html close;


/*Train Set*/
*Scoring for prediction on Train set;
data tree_train_score;
set traindata;
%include "C:\00\00-DATA SCIENCE\15-SAS_Project\SAS_FinalProject\treeoutput.sas"; 
run;

ods html on;
ods html style=journal;
data tree_train_score;
set tree_train_score;
Pred_HeartDisease = (P_HeartDisease1 >= 0.5); /* Convert probability to binary class */
run;
proc freq data=tree_train_score;
tables HeartDisease * Pred_HeartDisease / norow nocol nopercent;
run;
ods html close;


/*Test Set*/
*Scoring for prediction on Test set;
data tree_test_score;
set testData;
%include "C:\00\00-DATA SCIENCE\15-SAS_Project\SAS_FinalProject\treeoutput.sas";
run;

ods html on;
ods html style=journal;
data tree_test_score;
set tree_test_score;
Pred_HeartDisease = (P_HeartDisease1 >= 0.5); /* Convert probability to binary class */
run;
proc freq data=tree_test_score;
tables HeartDisease * Pred_HeartDisease / norow nocol nopercent;
run;
ods html close;

/*Test Set ROC Curve*/
ods html;
ods graphics on;
proc logistic data=tree_test_score plots=roc;
    model HeartDisease(event="1") = P_HeartDisease1 / nofit;
    roc 'Decision Tree' P_HeartDisease1;
    roccontrast;
run;
ods graphics off;





