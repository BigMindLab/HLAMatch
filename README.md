# HLAmatch
HLAmatch: A program that automates HLA compatibility evaluation across the five classic loci, classifying each into categories and generating an overall score (1–10). 

Contains 4 files: 
1. main.py
2. hla_finder.py (analyzes a PDF/Word document and exports the HLA data to a .csv file)
3. histomatch_analisis.py (analyzes the .csv file and exports the compatibility results to a .docx file named "Analisis_HLA")
4. histomatch_reporte.py (analyzes the .docx file and summarizes key information into another .docx file named "Reporte_HLA")


HLAmatch has 6 diagnostic possibilities for each locus:

1. Identical

Definition: Defined when there is a complete and absolute match between the compared alleles.
/
Condition: The number of differences is zero (len(dif) == 0) and at least one matching allele exists (len(inter) > 0).
/
Meaning: The profiles are exactly the same at that locus.

2. Included/
Definition: A special high-resolution state where one set of alleles is an almost complete subset of the other./
Condition: Activated when the overall match rate is low (less than 40%), but one of the allele sets is contained within the other by 90% or more (coverage >= 0.9)./
Meaning: Although not identical, the information in one set is virtually contained in the other.

3. Compatible
Definition: The default state when there are no severe conflicts, but full identity is not reached.
Condition: Assigned if the rules for "Identical", "Uncertain", or "Non-compatible" are not met. It is also assigned if high-resolution details are missing but the base groups match.
Meaning: Acceptable compatibility exists under the script's parameters.

4. Uncertain
Definition: Used when the data is ambiguous or the match rate is very low, but not zero.
Condition: Defined if the match rate is under 40% (porc_iguales < 0.4) with at least one matching allele (len(inter) >= 1), provided it does not qualify as "Included".
Meaning: Data suggests potential compatibility, but details are insufficient for confirmation.

5. Non-compatible
Definition: Defined when there is a clear discrepancy in the genetic structure.
Condition: Triggered if the Base Groups (the region before the colon :) are different, or if the match rate is under 40% with no shared alleles.
Meaning: Genetic rejection or lack of concordance exists at that locus.

6. No Data / Not Detected
Definition: States representing errors or missing information.
Condition: Activated if any group field is empty or if the mapping function cannot locate the DQA/DQB columns.
Meaning: Analysis could not be performed due to missing input parameters.
