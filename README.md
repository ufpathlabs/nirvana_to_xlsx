# nirvana_to_xlsx
Nirvana annotated Json file to xlsx format

This script was created by Johnny Bravo for UFPathLabs Molecular Labratory to create a summary .xlsx file for the review of NGS data from Illuminas DRAGEN Enrichment application. 

The excel file allows for the quick review of SNVs, SVs, CNVs, TMB, sample summary statisitcs, coverage metrics and any regions of low coverage. 

The Nirvana output from Dragen provides clinical-grade annotation of genomic variants, such as SNVs, MNVs, insertions, deletions, indels, STRs, SV, and CNVs. Useing a VCF as input. The output is a structured JSON representation of all annotations and sample information extracted from the VCF.

This script requires the following files to run as is. If you do not have one of these files, the code should be modifed so the file is not required. 

    hard-filtered.annotations.json.gz
    sv.annotations.json.gz
    cnv.annotations.json.gz
    seg.called.merged
    tmb.metrics.csv
    summary.csv
    qc-coverage-region-1_coverage_metrics.csv
    bqc-coverage-region-1_read_cov_report.bed
