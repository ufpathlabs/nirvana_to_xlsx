#!/usr/bin/env python3

import argparse
import json
import pandas as pd
import gzip
import os
import openpyxl
import xlsxwriter



def parse_sample_id_args():
    """Parse input arguments.

    The input sample id will be used to grab all required qc csv files and all
    the required json data files.

    Keyword arguments:

    -s, --SAMPLE_ID  -- sample id
    -p, --PATH -- path to files

    Return:

    args -- the parsed arguments
    """
    parser = argparse.ArgumentParser(
        description='sample id for json data files and qc files'
    )
    parser.add_argument(
        '-s',
        metavar='--SAMPLE_ID',
        type=str,
        help='the input sample id',
        required=True
    )
    parser.add_argument(
        '-p',
        metavar='--PATH',
        type=str,
        help='the path to all files',
        required=True
    )
    args = parser.parse_args()
    return args


def parse_transcripts(transcripts):
    """Parses the first transcript for the variant

    Each variant in a position has a list of transcripts. We are looping through
    the dictionary to see if any of the transcripts have our desired source,
    refseq, and then we parse the additional info we want.

    Keyword arguments:

    transcripts -- the transcript dictionary for the variant

    Return:

    transcript_name -- the name of the transcript
    source          -- the source of the data
    bio_type        -- the type of transcript
    hgnc            -- the hgnc value of the transcript
    hgvsc           -- the hgvsc value of the transcript
    hgvsp           -- the hgvsp value of the transcript
    """
    if transcripts != 'NA':
        for transcript in transcripts:
            try:
                source = transcript['source']
            except KeyError:
                source = 'NA'
            if source == 'RefSeq':
                try:
                    transcript_name = transcript['transcript']
                except KeyError:
                    transcript_name = 'NA'
                try:
                    bio_type = transcript['bioType']
                except KeyError:
                    bio_type = 'NA'
                try:
                    hgnc = transcript['hgnc']
                except KeyError:
                    hgnc = "NA"
                try:
                    hgvsc = transcript['hgvsc']
                except KeyError:
                    hgvsc = "NA"
                try:
                    hgvsp = transcript['hgvsp']
                except KeyError:
                    hgvsp = "NA"
                return (
                    transcript_name,
                    source,
                    bio_type,
                    hgnc,
                    hgvsc,
                    hgvsp
                )
            else:
                transcript_name = 'NA'
                source = 'NA'
                bio_type = 'NA'
                hgnc = 'NA'
                hgvsc = 'NA'
                hgvsp = 'NA'
    else:
        transcript_name = transcripts
        source = transcripts
        bio_type = transcripts
        hgnc = transcripts
        hgvsc = transcripts
        hgvsp = transcripts
    return (
        transcript_name,
        source,
        bio_type,
        hgnc,
        hgvsc,
        hgvsp
    )


def parse_clinvar(clinvar):
    """Parses the clinvar data for a variant

    Each variant has some data from ClinVar that we flatten here. Four values
    are returned as a tuple. If the values are not present then a value of NA
    is returned.

    Keyword arguments:

    clinvar -- the first clinvar dictionary for the variant

    Return:

    clinvar_id            -- the id value for the clinvar entry
    clinvar_review_status -- the review status for the clinvar entry
    clinvar_phenotypes    -- the listed phenotypes for the clinvar entry
    clinvar_significance  -- the significance in the clinvar entry
    """
    if clinvar != 'NA':
        try:
            clinvar_id = clinvar['id']
        except KeyError:
            clinvar_id = 'NA'
        try:
            clinvar_review_status = clinvar['reviewStatus']
        except KeyError:
            clinvar_review_status = 'NA'
        try:
            clinvar_phenotypes = clinvar['phenotypes']
        except KeyError:
            clinvar_phenotypes = 'NA'
        try:
            clinvar_significance = clinvar['significance']
        except KeyError:
            clinvar_significance = 'NA'
    else:
        clinvar_id = clinvar
        clinvar_review_status = clinvar
        clinvar_phenotypes = clinvar
        clinvar_significance = clinvar
    return (
        clinvar_id,
        clinvar_review_status,
        clinvar_phenotypes,
        clinvar_significance
    )


def parse_clingen(clingen):
    """Parses the clingen data for a SV variant

        Each structural variant has some data from ClinGen that we flatten here.
        Three values are returned as a tuple. If the values are not present then
        a value of NA is returned.

        Keyword arguments:

        clingen -- the first clingen dictionary for the variant

        Return:

        clingen_id              -- the id value for the clingen entry
        clingen_interperitation -- the interperitation for the clingen entry
        clingen_phenotypes      -- the listed phenotypes for the clingen entry
    """
    if clingen != 'NA':
        try:
            clingen_id = clingen['id']
        except KeyError:
            clingen_id = 'NA'
        try:
            clingen_interperitation = clingen['clinicalInterpretation']
        except KeyError:
            clingen_interperitation = 'NA'
        try:
            clingen_phenotypes = clingen['phenotypes']
        except KeyError:
            clingen_phenotypes = 'NA'
    else:
        clingen_id = clingen
        clingen_interperitation = clingen
        clingen_phenotypes = clingen
    return (
        clingen_id,
        clingen_interperitation,
        clingen_phenotypes
    )


def parse_variant(position):
    """Parse the variants for a given position

    Each variant contains many data points we want to extract, the main ones
    being the contig, start position, stop position, reference allele,
    altternate allele, and variant type. On top of these data points are
    additional ones from public datasets like gnomAD and clinical datasets like
    ClinVar and ClinGen. This function takes a position and parses each variant
    listed. Here we only take a small amount of the potential data for each
    variant, so it would be easy to add an additional `try, except` block and
    dictionary key, entry pair for a new piece of information as needed.

    Keyword arguments:

    position -- the position dictionary

    Return:

    variants -- a list of dictionaries, one for each variant parsed
    """
    variants = []
    for variant in position['variants']:
        contig = variant['chromosome']
        start = variant['begin']
        stop = variant['end']
        ref = variant['refAllele']
        alt = variant['altAllele']
        var_type = variant['variantType']
        try:
            hgvsg_vid = variant['hgvsg']
        except KeyError:
            try:
                hgvsg_vid = variant['vid']
            except KeyError:
                hgvsg_vid = 'NA'
        try:
            clinvar = parse_clinvar(clinvar=variant['clinvar'][0])
        except KeyError:
            clinvar = parse_clinvar('NA')
        try:
            dbsnp = variant['dbsnp'][0]
        except KeyError:
            dbsnp = 'NA'
        try:
            global_minor_allele_freq = variant['globalAllele']['globalMinorAlleleFrequency']
        except KeyError:
            global_minor_allele_freq = 'NA'
        try:
            gnomad = variant['gnomad']['allAf']
        except KeyError:
            gnomad = 'NA'
        try:
            onekg = variant['oneKg']['allAf']
        except KeyError:
            onekg = 'NA'
        try:
            revel = variant['revel']['score']
        except KeyError:
            revel = 'NA'
        try:
            topmed = variant['topmed']['allAf']
        except KeyError:
            topmed = 'NA'
        try:
            transcript = parse_transcripts(transcripts=variant['transcripts'])
        except KeyError:
            transcript = parse_transcripts(transcripts='NA')
        variants.append({
            'Contig': contig,
            'Start': start,
            'Stop': stop,
            'Ref Allele': ref,
            'Alt Allele': alt,
            'Variant Type': var_type,
            'HGVSG/ VID': hgvsg_vid,
            'Clinvar ID': clinvar[0],
            'Clinvar Review Status': clinvar[1],
            'Clinvar Phenotypes': clinvar[2],
            'Clinvar Significance': clinvar[3],
            'dbSNP': dbsnp,
            'Global Minor Allele Freq': global_minor_allele_freq,
            'gnomAD': gnomad,
            'oneKG': onekg,
            'REVEL': revel,
            'topMED': topmed,
            'Transcript': transcript[0],
            'Transcript Source': transcript[1],
            'Transcript Bio Type': transcript[2],
            'Transcript HGNC': transcript[3],
            'Transcript HGVSC': transcript[4],
            'Transcript HGVSP': transcript[5]
        })
    return variants


def parse_position(data, var_type):
    """Parse the positions for a sample

    Each position contains some basic information like sample info, and position
    filtering, but they also contain lists for the variants, and some other
    quality metrics. This function parses each position for an input json file
    and returns a list of dictionaries, one dictionary for each position.

    Keyword arguments:

    data     -- the loaded json file/ data
    var_type -- 'SV' or 'SNV'

    Return:

    positions -- a list of dictionaries, one for each position parsed
    """
    positions = []
    for position in data['positions']:
        sample = position['samples'][0]
        try:
            filter = position['filters']
        except KeyError:
            filter = 'NA'
        if var_type == 'SNV':
            try:
                mapping_quality = position['mappingQuality']
            except KeyError:
                mapping_quality = 'NA'
            try:
                variant_freq = sample['variantFrequencies'][0]
            except KeyError:
                variant_freq = 'NA'
            try:
                total_depth = sample['totalDepth']
            except KeyError:
                total_depth = 'NA'
            try:
                allele_depths = sample['alleleDepths']
            except KeyError:
                allele_depths = 'NA'
            try:
                somatic_quality = sample['somaticQuality']
            except KeyError:
                somatic_quality = 'NA'
            variants = parse_variant(position=position)
            positions.append({
                'Filter': filter,
                'Mapping Quality': mapping_quality,
                'Variant Frequency': variant_freq,
                'Total Depth': total_depth,
                'Allele Depths': allele_depths,
                'Somatic Quality': somatic_quality,
                'Variants': variants
            })
        elif var_type == 'SV':
            try:
                split_read_counts = sample['splitReadCounts']
            except KeyError:
                split_read_counts = 'NA'
            try:
                paired_end_read_counts = sample['pairedEndReadCounts']
            except KeyError:
                paired_end_read_counts = 'NA'
            try:
                clingen = parse_clingen(position['clingen'][0])
            except KeyError:
                clingen = parse_clingen('NA')
            variants = parse_variant(position=position)
            positions.append({
                'Filter': filter,
                'Split Read Counts': split_read_counts,
                'Paired End Read Counts': paired_end_read_counts,
                'Clingen ID': clingen[0],
                'Clingen Interperitation': clingen[1],
                'Clingen Phenotypes': clingen[2],
                'Variants': variants
            })
    return positions


def parse_hits(positions, var_type):
    """Take the parsed positions and create flattened 'hits'

    This function takes the parsed positions from `parse_positions()` and
    creates a new list called hits. These hits are just the 'flattened' data
    from the positions and variants within those positions. This function is
    basically just cleaning up the data so it can be made into a pandas
    DataFrame with little effort.

    Keyword arguments:

    positions -- the list of parsed positions from `parse_positions()`
    var_type  -- 'SV' or 'SNV'

    Return:

    df -- the pd DataFrame of the positions
    """
    hits = []
    for position in positions:
        for variant in position['Variants']:
            var_dict = {
                'Contig': variant['Contig'],
                'Start': variant['Start'],
                'Stop': variant['Stop'],
                'Ref Allele': variant['Ref Allele'],
                'Alt Allele': variant['Alt Allele'],
                'Variant Type': variant['Variant Type'],
                'Transcript': variant['Transcript'],
                'Transcript Source': variant['Transcript Source'],
                'Transcript Bio Type': variant['Transcript Bio Type'],
                'Transcript HGNC': variant['Transcript HGNC'],
                'Transcript HGVSC': variant['Transcript HGVSC'],
                'Transcript HGVSP': variant['Transcript HGVSP'],
                'HGVSG/ VID': variant['HGVSG/ VID'],
                'Filter': position['Filter'],
                'Clinvar ID': variant['Clinvar ID'],
                'Clinvar Review Status': variant['Clinvar Review Status'],
                'Clinvar Phenotypes': variant['Clinvar Phenotypes'],
                'Clinvar Significance': variant['Clinvar Significance'],
                'dbSNP': variant['dbSNP'],
                'Global Minor Allele Freq': variant['Global Minor Allele Freq'],
                'gnomAD': variant['gnomAD'],
                'oneKG': variant['oneKG']
            }
            if var_type == 'SNV':
                var_dict['Mapping Quality'] = position['Mapping Quality']
                var_dict['Variant Frequency'] = position['Variant Frequency']
                var_dict['Total Depth'] = position['Total Depth']
                var_dict['Allele Depths'] = position['Allele Depths']
                var_dict['Somatic Quality'] = position['Somatic Quality']
                var_dict['REVEL'] = variant['REVEL']
                var_dict['topMED'] = variant['topMED']
            elif var_type == 'SV':
                var_dict['Split Read Counts'] = position['Split Read Counts']
                var_dict['Paired End Read Counts'] = position['Paired End Read Counts']
                var_dict['Clingen ID'] = position['Clingen ID']
                var_dict['Clingen Interperitation'] = position['Clingen Interperitation']
                var_dict['Clingen Phenotypes'] = position['Clingen Phenotypes']
            hits.append(var_dict)
    df = pd.DataFrame(hits)
    return df


def parse_json(json_file, var_type):
    """Load in the json file and parse the data

    This function is rather simple. It is loading in the gzipped json data,
    parsing the positions for the input file in `parse_positions()`, and then
    flattening the data in `parse_hits()`. At the end the returned value is the
    same DataFrame that is returned from `parse_hits()`.

    Keyword arguments:

    json_file -- the sv or snv json file to be parsed
    var_type  -- 'SV' or 'SNV'

    Return:

    df -- the DataFrame created in `parse_hits()`
    """
    with gzip.open(json_file, 'r') as j:
        data = json.load(j)
        positions = parse_position(data=data, var_type=var_type)
    df = parse_hits(positions=positions, var_type=var_type)
    return df


def parse_qc_csv(csv_file, qc_type):
    """Parse the qc metrics csv files

    This function handles reading in all the extra qc csv files. The three
    options are the tmb, summary, or coverage csv file.

    Keyword arguments:

    csv_file -- the input qc csv file to be parsed
    qc_type  -- 'TMB', 'SUMMARY', or 'COVERAGE'

    Return:

    df -- a DataFrame of the input csv file
    """
    if qc_type == 'TMB':
        df = pd.read_csv(
            csv_file,
            usecols=[2, 3],
            names=['TMB Summary', 'Value']
        )
    elif qc_type == 'SUMMARY':
        df = pd.read_csv(
            csv_file,
            skiprows=4,
            names=['DRAGEN Enrichment Summary Report', 'Value']
        )
    elif qc_type == 'COVERAGE':
        df = pd.read_csv(
            csv_file,
            usecols=[2, 3],
            names=['Coverage Summary', 'Value']
        )
    return df


def parse_bed(bed_file):
    """Parse the input bed file

    This function parses the input bed file and returns a DataFrame containing
    only the entries from the bed file where the total coverage for the region
    is at or below 500.

    Keyword arguments:

    bed_file -- the input bed file

    Return:

    df -- a DataFrame of the low coverage regions from the bed file
    """
    low_cov_regions = []
    with open(bed_file, 'r') as bed:
        lines = bed.readlines()[1:]
        for line in lines:
            row = line.split('\t')
            if int(row[5]) <= 500:
                low_cov_regions.append({
                    'Contig': row[0],
                    'Start': row[1],
                    'End': row[2],
                    'Name': row[3],
                    'Gene ID': row[4],
                    'Total Coverage': row[5],
                    'Read1 Coverage': row[6],
                    'Read2 Coverage': row[7]
                })
    df = pd.DataFrame(low_cov_regions)
    return df


def parse_cnv_seg(cnv_seg):
    """Pasing the seg file containing CNVs

    This function breaks down the seg file into individual dictionaries so
    that it can be merged with the data from the CNV json file.

    Keyword arguments:

    cnv_seg -- the input cnv seg file

    Return:

    cnvs -- a list of dictionaries from the seg file
    """
    cnvs = []
    with open(cnv_seg, 'r') as seg:
        next(seg)
        for line in seg:
            entry = line.split('\t')
            sample = entry[0]
            chrom = entry[1]
            start = int(entry[2])
            end = int(entry[3])
            num_targets = entry[4]
            seg_mean = entry[5]
            seg_call = entry[6]
            qual = entry[7]
            qual_filter = entry[8]
            copy_number = entry[9]
            ploidy = entry[10]
            imp_pairs = entry[11].strip()
            cnvs.append({
                'Sample': sample,
                'Chromosome': chrom,
                'Start': start,
                'End': end,
                'Num_Targets': num_targets,
                'Segment_Mean': seg_mean,
                'Segment_Call': seg_call,
                'Qual': qual,
                'Filter': qual_filter,
                'Copy_Number': copy_number,
                'Ploidy': ploidy,
                'Improper_Pairs': imp_pairs
            })
    return cnvs


def parse_cnv_json(cnv_json):
    """Pasing the json file containing CNVs

    This function breaks down the json file into individual dictionaries so
    that it can be merged with the data from the CNV seg file.

    Keyword arguments:

    cnv_json -- the input cnv json file

    Return:

    cnvs -- a list of dictionaries from the json file
    """
    cnvs = []
    with gzip.open(cnv_json, 'r') as j:
        data = json.load(j)
        for position in data['positions']:
            sv_len = position['svLength']
            cytoband = position['cytogeneticBand']
            for variant in position['variants']:
                contig = variant['chromosome']
                start = variant['begin']
                stop = variant['end']
                var_type = variant['variantType']
                try:
                    for transcript in variant['transcripts']:
                        if transcript['source'] == 'RefSeq':
                            source = transcript['source']
                            try:
                                transcript_name = transcript['transcript']
                            except KeyError:
                                transcript_name = 'NA'
                            try:
                                bio_type = transcript['bioType']
                            except KeyError:
                                bio_type = 'NA'
                            try:
                                hgnc = transcript['hgnc']
                            except KeyError:
                                hgnc = "NA"
                            break
                    if source != 'RefSeq':
                        source = variant['transcripts'][0]['source']
                        try:
                            transcript_name = variant['transcripts'][0]['transcript']
                        except KeyError:
                            transcript_name = 'NA'
                        try:
                            bio_type = variant['transcripts'][0]['bioType']
                        except KeyError:
                            bio_type = 'NA'
                        try:
                            hgnc = variant['transcripts'][0]['hgnc']
                        except KeyError:
                            hgnc = "NA"
                except KeyError:
                    transcript_name = 'NA'
                    bio_type = 'NA'
                    hgnc = 'NA'
            cnvs.append({
                'Chromosome': contig,
                'Start': int(start),
                'End': int(stop),
                'svLength': sv_len,
                'cytogeneticBand': cytoband,
                'variantType': var_type,
                'transcript': transcript_name,
                'bioType': bio_type,
                'hgnc': hgnc
            })
    return cnvs


def parse_cnvs(seg_cnvs, json_cnvs):
    """Merging the seg and json CNV files

    This function creates DataFrames from the seg and json lists of dictionaries
    and then merges them on their contigs, starts, and stops.

    Keyword arguments:

    seg_cnvs  -- the list of dictionaries returned from `parse_cnv_seg()`
    json_cnvs -- the list of dictionaries returned from `parse_cnv_json()`

    Return:

    cnvs_df -- the merged DataFrame of the seg and json CNVs
    """
    seg_df = pd.DataFrame(seg_cnvs)
    seg_df['Start'] = seg_df['Start'] + 1
    json_df = pd.DataFrame(json_cnvs)
    cnvs_df = pd.merge(seg_df, json_df, how='outer', on=['Chromosome', 'Start', 'End'])

    return cnvs_df

def add_tmb(tmb_trace, snv_hits):
    """This function creates a DataFrame from the tmb trace file and then merges with SNV hits based on
    their chromosome, starts, and stops. The tmb trace file returns the data decision for marking a variant for
    TMB counts.

    Keyword arguments: tmb_trace, snv_hits

    Return: snv_tmb_merge -- the merged DataFrame of the snv hits and the tmb trace data
    """
    #pull in the trace tmb file and use cols listed
    tmb_df = pd.read_csv(tmb_trace, sep='\t', usecols= ['Chromosome','Position','RefAllele','AltAllele','VAF',
                                                        'VariantType','TotalDepth','DbMaxAlleleCount',
                                                        'CosmicMaxCount','withinValidTmbRegion','tmbCandidate',
                                                        'nonsyn','databaseFilter','proxiFilter','tmbVariant',
                                                        'withinValidNonsynRegion'])
    #rename the tmbdf to match teh snv df
    tmb_df.rename(columns={"Chromosome": "Contig", "Position": "Start", "RefAllele": "Ref Allele"}, inplace=True)

    #pull in the snv df want to add the trace tmb data to
    snv_df = pd.DataFrame(snv_hits)

    #match data types
    tmb_df['Contig'] = tmb_df['Contig'].astype(object)
    tmb_df['Start'] = pd.to_numeric(tmb_df['Start'].replace(',','', regex=True), errors='coerce')
    tmb_df['Start'] = tmb_df['Start'].astype('int64', errors='ignore')

    #if VatriatType = deletion/indel then + 1 to the position
    pd.options.mode.chained_assignment = None
    tmb_df['Start'][(tmb_df['VariantType'] == 'INSERTION')] = tmb_df['Start'] + 1
    tmb_df['Start'][(tmb_df['VariantType'] == 'DELETION')] = tmb_df['Start'] + 1
    snv_df['Start'][(snv_df['Variant Type'] == 'indel')] = snv_df['Start'] + 1

    #merge the datframes
    snv_tmb_merge = pd.merge(snv_df, tmb_df, how='outer', on=['Contig', 'Start'])
    snv_tmb_merge = snv_tmb_merge.sort_values(by=['tmbVariant','nonsyn'], ascending=False)
    snv_tmb_merge.fillna(0, inplace=True)

    return snv_tmb_merge

def get_snp_id(snv_hits):
    """
    """
    snps = ('rs1005533', 'rs1024116', 'rs1028528', 'rs10495407', 'rs10771010', 'rs11781516', 'rs13050660', 'rs1335873',
            'rs1357617', 'rs1360288', 'rs136337', 'rs1382387', 'rs1413212', 'rs1454361', 'rs1463729', 'rs1468118',
            'rs1493232', 'rs1982986', 'rs1994997', 'rs2010253', 'rs2040411', 'rs2046361', 'rs2056277', 'rs2076848',
            'rs214054', 'rs2247221', 'rs2518968', 'rs251934', 'rs2714854', 'rs2831700', 'rs354439', 'rs3819854',
            'rs717302', 'rs727811', 'rs729172', 'rs740910', 'rs8037429', 'rs826472', 'rs876724', 'rs891700', 'rs901398',
            'rs914165', 'rs9583190', 'rs964681')
    snp_id_hits = snv_hits[snv_hits['dbSNP'].isin(snps)]
    return snp_id_hits

def parse_cnv_filtered(cnv_hits):
    """Passing the merged seg and json CNV dataframe

        This functions uses the dataframe from the merged seg and json files to create a new
        dataframe. This function removes data from the merged seg and json CNV dataframe that
        does NOT meet the following requirements:

        Segment_Mean >= 2.5 OR <=0.5
        biotype != pseudogene
        Filter = PASS

        Keyword arguments:

        cnv_hits -- the merged DataFrame of the seg and json CNVs

        Return:

        cnv_filtered -- the merged DataFrame of the seg and json CNVs filtered with the
        given conditions.

        Added by NYL
        """

    cnv_filtered = cnv_hits[(cnv_hits.bioType != 'pseudogene') & (cnv_hits.Filter == 'PASS') &
                            ((cnv_hits.Segment_Mean.astype(float) < 0.65) |
                             (cnv_hits.Segment_Mean.astype(float) > 1.9))]

    # not sure if i want this or not. Will limit hits to high qual.
    # cnv_filtered = cnv_filtered[(cnv_filtered['Qual'].astype(int) > 190) &
    #                             (cnv_filtered['Segment_Mean'].astype(float) > 0.6) |
    #                             (cnv_filtered['Segment_Mean'].astype(float) < 0.6) |
    #                             (cnv_filtered['Segment_Mean'].astype(float) > 1.9)]

    # keep genes in list
    #bed_files_dir = '/Volumes/files/PATH/DRL/Molecular/NGS21/Bioinformatic_Pipelines/SOPs_and_Version_Documentation/Files_to_run_localy/'
    bed_files_dir = '/ext/path/DRL/Molecular/NGS21/Bioinformatic_Pipelines/SOPs_and_Version_Documentation/Files_to_run_localy/'
    CNV_Filtered_Bed_File = pd.read_csv(f'{bed_files_dir}CNV_Filtered_Bed_File.csv')
    CNV_Filtered_Bed_File_Keep = pd.read_csv(f'{bed_files_dir}CNV_Filtered_Bed_File.csv')

    # checks to see if cnv calls are in the filtered gene list and keeps only rows that are in the filtered gene list
    with pd.option_context("mode.chained_assignment", None):
        for i, r in cnv_filtered.iterrows():
            # noinspection PyTypeChecker
            cnv_filtered.loc[i, 'in_filtered_gene_list'] = \
                any((r.Chromosome == CNV_Filtered_Bed_File.Chromosome) & (r.Start > CNV_Filtered_Bed_File.Start) &
                    (r.Start < CNV_Filtered_Bed_File.End)) | \
                any((r.Chromosome == CNV_Filtered_Bed_File.Chromosome) &
                    (r.End > CNV_Filtered_Bed_File.Start) & (r.End < CNV_Filtered_Bed_File.End)) | \
                any((r.Chromosome == CNV_Filtered_Bed_File.Chromosome) & (r.Start < CNV_Filtered_Bed_File.Start) &
                    (r.End > CNV_Filtered_Bed_File.Start)) | \
                any((r.Chromosome == CNV_Filtered_Bed_File.Chromosome) & (r.Start > CNV_Filtered_Bed_File.End) &
                    (r.End < CNV_Filtered_Bed_File.End))
            if cnv_filtered.empty != True:
                continue
            else:
                cnv_filtered = cnv_filtered[(cnv_filtered['in_filtered_gene_list'] == True)]

    # filters the non-supprssor genes based on the more strict criteria (keeps <0.6 for non-suppressor and <0.75  if suppressor)
    # and if > 1.9 for the oncogenes.
    # changing true false in the sheet will define if reportable as an amp or del. If suppressor =TRUE then only report
    # as a deletion. if oncogene = TRUE then only report as AMP if both are true, will report both.

    suppressor_genes = CNV_Filtered_Bed_File_Keep.loc[CNV_Filtered_Bed_File_Keep['suppressor_gene'] == True]
    Oncogenes = CNV_Filtered_Bed_File_Keep.loc[CNV_Filtered_Bed_File_Keep['Oncogene'] == True]
    
    with pd.option_context("mode.chained_assignment", None):
        for i, r in cnv_filtered.iterrows():
            # noinspection PyTypeChecker
            cnv_filtered.loc[i, 'suppressor_gene'] = \
                any((r.Chromosome == suppressor_genes.Chromosome) & (r.Start > suppressor_genes.Start) &
                    (r.Start < suppressor_genes.End)) | \
                any((r.Chromosome == suppressor_genes.Chromosome) &
                    (r.End > suppressor_genes.Start) & (r.End < suppressor_genes.End)) | \
                any((r.Chromosome == suppressor_genes.Chromosome) & (r.Start < suppressor_genes.Start) &
                    (r.End > suppressor_genes.Start)) | \
                any((r.Chromosome == suppressor_genes.Chromosome) & (r.Start > suppressor_genes.End) &
                    (r.End < suppressor_genes.End))

    with pd.option_context("mode.chained_assignment", None):
        for i, r in cnv_filtered.iterrows():
            # noinspection PyTypeChecker
            cnv_filtered.loc[i, 'Oncogene'] = \
                any((r.Chromosome == Oncogenes.Chromosome) & (r.Start > Oncogenes.Start) &
                    (r.Start < Oncogenes.End)) | \
                any((r.Chromosome == Oncogenes.Chromosome) &
                    (r.End > Oncogenes.Start) & (r.End < Oncogenes.End)) | \
                any((r.Chromosome == Oncogenes.Chromosome) & (r.Start < Oncogenes.Start) &
                    (r.End > Oncogenes.Start)) | \
                any((r.Chromosome == Oncogenes.Chromosome) & (r.Start > Oncogenes.End) &
                    (r.End < Oncogenes.End))

    with pd.option_context("mode.chained_assignment", None):
        cnv_filtered['Segment_Mean'] = pd.to_numeric(cnv_filtered['Segment_Mean'], errors='coerce')

    cnv_filtered = cnv_filtered[(cnv_filtered['suppressor_gene'].astype(str) == 'True') &
                                (cnv_filtered['Segment_Mean'].astype(float) < 1) |
                                (cnv_filtered['Oncogene'].astype(str) == 'True') &
                                (cnv_filtered['Segment_Mean'].astype(float) > 1)]

    # creates a new bed file with the segment hits so this can me merged with the segemnts so
    # we can add the gene names for each region
    with pd.option_context("mode.chained_assignment", None):
        for i, r in CNV_Filtered_Bed_File.iterrows():
            # noinspection PyTypeChecker
            CNV_Filtered_Bed_File.loc[i, 'in_cnv_filtered'] = \
                any((r.Chromosome == cnv_filtered.Chromosome) & (r.Start > cnv_filtered.Start) & (
                            r.Start < cnv_filtered.End)) | \
                any((r.Chromosome == cnv_filtered.Chromosome) & (r.End > cnv_filtered.Start) & (
                            r.End < cnv_filtered.End))

    # keeps the rows in the filtered bed file
    CNV_Filtered_Bed_File = CNV_Filtered_Bed_File[(CNV_Filtered_Bed_File['in_cnv_filtered'] == True)]

    # merges the two together so we can have the gene names and sorts the data so the gene names show up near the correct region
    cnv_filtered = cnv_filtered.append(CNV_Filtered_Bed_File)
    cnv_filtered = cnv_filtered.sort_values(by=['Sample']).fillna('').drop(
        columns=['in_filtered_gene_list', 'in_cnv_filtered'])
    cnv_filtered["Genes"] = cnv_filtered['hgnc'].astype(str) + cnv_filtered['hgnc_symbol']
    cnv_filtered['Sum'] = cnv_filtered['Start'] + cnv_filtered['End']
    cnv_filtered = cnv_filtered.sort_values(by=['Chromosome', 'Sum', 'Start'], ascending=[True, True, False])
    cnv_filtered['Reportable_gene'] = cnv_filtered['Genes'].isin(CNV_Filtered_Bed_File_Keep['hgnc_symbol'])

    # data to display
    cnv_filtered = cnv_filtered[
        ['Sample', 'Chromosome', 'Start', 'End', 'Num_Targets', 'Segment_Mean', 'suppressor_gene', 'Oncogene',
         'cytogeneticBand', 'Genes', 'Reportable_gene', 'Qual', 'variantType', 'Segment_Call', 'Filter',
         'Copy_Number', 'Ploidy', 'Improper_Pairs','svLength', 'transcript', 'bioType']]

    return cnv_filtered

def out_corrected_cnv_filtered(cnv_filter):
    """This function adds the data needed for the excel sheet but doesnt mess up the write updated cnv vcf function"""
    out_correct_cnv = cnv_filter.mask(cnv_filter == '')

# make sure empty segment mean gets filled so it can be uploaded as vcf
    out_correct_cnv['Segment_Mean'] = out_correct_cnv.groupby(['Chromosome', 'Genes'], sort=False)['Segment_Mean'].apply(lambda x: x.ffill().bfill())
    out_correct_cnv['Segment_Mean'] = out_correct_cnv.groupby(['Chromosome'], sort=False)['Segment_Mean'].apply(lambda x: x.ffill().bfill())
    #out_correct_cnv['svLength'] = out_correct_cnv.groupby(['Chromosome'], sort=False)['svLength'].apply(lambda x: x.ffill().bfill())
    out_correct_cnv['variantType'] = out_correct_cnv.groupby(['Chromosome', 'Genes'], sort=False)['variantType'].apply(lambda x: x.ffill().bfill())
    out_correct_cnv['variantType'] = out_correct_cnv.groupby(['Chromosome'], sort=False)['variantType'].apply(lambda x: x.ffill().bfill())
    #out_correct_cnv['Improper_Pairs'] = out_correct_cnv.groupby(['Chromosome'], sort=False)['Improper_Pairs'].apply(lambda x: x.ffill().bfill())
    #out_correct_cnv['Qual'] = out_correct_cnv.groupby(['Chromosome', 'Genes'], sort=False)['Qual'].apply(lambda x: x.ffill().bfill())
    #out_correct_cnv['Num_Targets'] = out_correct_cnv.groupby(['Chromosome', 'Genes'], sort=False)['Num_Targets'].apply(lambda x: x.ffill().bfill())
    #out_correct_cnv['Filter'] = out_correct_cnv.groupby(['Chromosome'], sort=False)['Filter'].apply(lambda x: x.ffill().bfill())

    # filter out any calls not following suppressor/oncogene ruls
    out_correct_cnv = out_correct_cnv[(out_correct_cnv['suppressor_gene'].astype(str) == 'True') &
                                    (out_correct_cnv['Segment_Mean'].astype(float) < 1) |
                                    (out_correct_cnv['Oncogene'].astype(str) == 'True') &
                                    (out_correct_cnv['Segment_Mean'].astype(float) > 1)]
    return out_correct_cnv


def write_updated_cnv_vcf(cnv_filter, cnv_vcf):
    """this function writes the cnv outfile based on the filtered calls"""

# turn all empty cells to nan so can use fillna later
    short_cnv_list = cnv_filter.mask(cnv_filter == '')

#make sure empty segment mean gets filled so it can be uploaded as vcf
    short_cnv_list['Segment_Mean'] = short_cnv_list.groupby(['Chromosome','Genes'], sort=False)['Segment_Mean'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['Segment_Mean'] = short_cnv_list.groupby(['Chromosome'], sort=False)['Segment_Mean'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['svLength'] = short_cnv_list.groupby(['Chromosome',], sort=False)['svLength'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['variantType'] = short_cnv_list.groupby(['Chromosome', 'Genes'], sort=False)['variantType'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['variantType'] = short_cnv_list.groupby(['Chromosome'], sort=False)['variantType'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['Improper_Pairs'] = short_cnv_list.groupby(['Chromosome'], sort=False)['Improper_Pairs'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['Qual'] = short_cnv_list.groupby(['Chromosome'], sort=False)['Qual'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['Num_Targets'] = short_cnv_list.groupby(['Chromosome'], sort=False)['Num_Targets'].apply(lambda x: x.ffill().bfill())
    short_cnv_list['Filter'] = short_cnv_list.groupby(['Chromosome'], sort=False)['Filter'].apply(lambda x: x.ffill().bfill())

# find start and end for genes with multiple calls and merge segment means
    start_min_value = short_cnv_list.groupby(['Genes','cytogeneticBand']).Start.min().to_frame()
    end_max_value = short_cnv_list.groupby(['Genes','cytogeneticBand']).End.max().to_frame()
    seg_mean_average = short_cnv_list.groupby(['Genes','cytogeneticBand']).Segment_Mean.mean().to_frame()
    update_data = pd.merge(pd.merge(start_min_value, end_max_value, on=['Genes','cytogeneticBand']),seg_mean_average,on=['Genes','cytogeneticBand'])

#add in min max start and merged segment mean for merged genes

    short_cnv_list = short_cnv_list.merge(update_data, on=['Genes','cytogeneticBand'], how='outer')
    short_cnv_list['Start_y'] = short_cnv_list['Start_y'].fillna(short_cnv_list['Start_x'])
    short_cnv_list['End_y'] = short_cnv_list['End_y'].fillna(short_cnv_list['End_x'])
    short_cnv_list['Segment_Mean_y'] = short_cnv_list['Segment_Mean_y'].fillna(short_cnv_list['Segment_Mean_x'])
    short_cnv_list = short_cnv_list.drop(['Start_x','End_x','Segment_Mean_x'], axis=1)
    short_cnv_list = short_cnv_list.rename(columns={"Start_y":"Start","End_y":"End","Segment_Mean_y":"Segment_Mean"})

#update Start and end to numeric
    short_cnv_list.Start = pd.to_numeric(short_cnv_list.Start, errors='coerce')
    short_cnv_list.End = pd.to_numeric(short_cnv_list.End, errors='coerce')

#keep only unique records for vcf upload. keep highest qual if there is a duplicate

    short_cnv_list.Qual = pd.to_numeric(short_cnv_list.Qual, errors='coerce')
    #short_cnv_list = short_cnv_list.sort_values('Qual', ascending=False).drop_duplicates('Genes', keep='first')
    short_cnv_list = short_cnv_list.sort_values('cytogeneticBand', ascending=False).drop_duplicates('Genes', keep='last')
    short_cnv_list.drop(short_cnv_list[short_cnv_list['Reportable_gene'] == False].index, inplace=True)

#filter out any calls not following suppressor/oncogene ruls
    short_cnv_list = short_cnv_list[(short_cnv_list['suppressor_gene'].astype(str) == 'True') &
                                    (short_cnv_list['Segment_Mean'].astype(float) < 1) |
                                    (short_cnv_list['Oncogene'].astype(str) == 'True') &
                                    (short_cnv_list['Segment_Mean'].astype(float) > 1)]

#add columns for vcf info and turn svLenght to intiger so it will load to qci
    ID2_mapping = {'copy_number_loss': 'LOSS', 'copy_number_gain': 'GAIN'}
    ALT_mapping = {'copy_number_loss': '<DEL>', 'copy_number_gain': '<DUP>'}
    with pd.option_context('mode.chained_assignment', None):
        short_cnv_list['ID2'] = short_cnv_list['variantType'].replace(ID2_mapping)
        short_cnv_list['ALT'] = short_cnv_list['variantType'].replace(ALT_mapping)
        short_cnv_list['svLength'] = short_cnv_list['svLength'].astype(int)
        short_cnv_list['Start'] = short_cnv_list['Start'].astype(int)
        short_cnv_list['End'] = short_cnv_list['End'].astype(int)
        short_cnv_list['Length'] = (short_cnv_list['Start'] - short_cnv_list['End']).abs()

#pulls the cnvs and formats vcf style for qci upload
#CHROM	POS	ID	REF	ALT	QUAL	FILTER	INFO	FORMAT
#chr10	89623194	DRAGEN:LOSS:chr10:89623194-89728532	N	<DEL>	200	PASS	END=89728532;REFLEN=1152500;SVTYPE=CNV;SVLEN=1152500	GT:FC:BC:PE	./.:0.33224369780928004:20:50,1

    cnvs = short_cnv_list.apply(lambda x: f"{x['Chromosome']}\t{x['Start']}\t DRAGEN:{x['ID2']}:{x['Chromosome']}:"
                                          f"{x['Start']}-{x['End']}\tN\t{x['ALT']}\t{x['Qual']}\t{x['Filter']}\t"
                                          f"END={x['End']};REFLEN={x['Length']};SVTYPE=CNV;SVLEN={x['Length']}"
                                          f"\tGT:FC:BC:PE\t./.:{x['Segment_Mean']}:{x['Num_Targets']}:{x['Improper_Pairs']}", axis=1)


#get the CNV.VCF header
    with gzip.open(cnv_vcf, 'rt') as cnvfile, open(f'{cnv_vcf}_FILTERED.CNV.VCF', 'w') as CNV_header:
        for line in cnvfile:
            if '#' in line:
                CNV_header.write(line)

#replace SM with FC
    with open(f'{cnv_vcf}_FILTERED.CNV.VCF', 'r') as update_header:
        update = update_header.read()
        update = update.replace('SM','FC')
    with open(f'{cnv_vcf}_FILTERED.CNV.VCF', 'w') as update_header:
        update_header.write(update)

# append the variants
    with open(f'{cnv_vcf}_FILTERED.CNV.VCF', 'a') as out_cnv_vcf:
        for cnv in cnvs:
            if not cnvs.empty:
                out_cnv_vcf.write(cnv)
                out_cnv_vcf.write('\n')


def write_xlsx(data, sample_id):
    """Write the output xlsx file

    This function takes all the data generated from parsing the sv and snv
    json files, the tmb, summary, and coverage csv files, and the bed file.
    Each DataFrame is written to its own sheet, or tab, within an excel file/
    workbook.

    Keyword arguments:

    data      -- a tuple containing all the parsed file's data
    sample_id -- the sample id
    """
    with pd.ExcelWriter(f'{sample_id}.xlsx') as writer:
        data[0].to_excel(writer, sheet_name='SNVs', index=False)
        data[1].to_excel(writer, sheet_name='SNP_ID', index=False)
        data[2].to_excel(writer, sheet_name='SVs', index=False)
        data[3].to_excel(writer, sheet_name='CNVs', index=False)
        data[4].to_excel(writer, sheet_name='CNVs Filtered', index=False)
        data[5].to_excel(writer, sheet_name='TMB', index=False)
        data[6].to_excel(writer, sheet_name='TMB_FILTER', index=False)
        data[7].to_excel(writer, sheet_name='SUMMARY', index=False)
        data[8].to_excel(writer, sheet_name='COVERAGE', index=False)
        data[9].to_excel(writer, sheet_name='LOW COVERAGE', index=False)

        # filter the TMB_Filter tab to only show tmb variants
        # Get the xlsxwriter worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['TMB_FILTER']
        # Get the dimensions of the dataframe.
        (max_row, max_col) = data[6].shape


        # Set the autofilter
        worksheet.autofilter(0, 0, max_row, max_col - 1)

        # Add filter criteria
        worksheet.filter_column('AP', 'tmbVariant == 1')
        worksheet.filter_column('AM', 'nonsyn == 1')
        worksheet.filter_column('N', "Filter == ['PASS']")

        # It isn't enough to just apply the criteria. The rows that don't match
        # must also be hidden. We use Pandas to figure our which rows to hide.
        #need to reset the index for this to function
        data[6].reset_index(drop=True, inplace=True)

        #turns cells green
        format1 = workbook.add_format({'bg_color': '#CCEECE', 'font_color': '#225F00'})

        #this adds lines around all squares if want to add
        #border_fmt = workbook.add_format({'bottom': 1, 'top': 1, 'left': 1, 'right': 1})

        for r in (data[6].index[(data[6]['nonsyn'] == 1)].tolist()):
            worksheet.set_row(r + 1, None, format1)

        for row_num in (data[6].index[(data[6]['tmbVariant'] != 1)].tolist()):
            worksheet.set_row(row_num + 1, options={'hidden': True})

        # if want to hide nonsyn instead of highlighting tmbvariants, use this
        # for row_num in (data[6].index[(data[6]['nonsyn'] != 1)].tolist()):
        #     worksheet.set_row(row_num + 1, options={'hidden': True})


def write_xlsx_cnv_filtered(data, sample_id):
    """Write the output xlsx file

    This function takes all the data generated from the merged seg and json files and creates
    a second, separate Excel file and writes in the cnv_filter DataFrame.

    Keyword arguments:

    data      -- CNV filtered dataframe
    sample_id -- the sample id
    """

#cant get .xlsx to work try .csv
    exists = os.path.isfile('CNV-filtered_SUMMARY.csv')
    if not exists:
        data[0].to_csv('CNV-filtered_SUMMARY.csv', encoding='utf-8', index=False)
    else:
        data[0].to_csv('CNV-filtered_SUMMARY.csv', mode='a', index=False, header=False)



def main():
    """Main function that runs

    1. Parse the arguments to get the sample id as `args.s`
    2. Parse the SNV json file; var_type = SNV
    3. Parse the SV json file; var_type = SV
    4. Parse the CNV json file; var_type = CNV
    5. Parse the TMB csv file; qc_type = TMB
    6. Parse the SUMMARY csv file; qc_type = SUMMARY
    7. Parse the COVERAGE csv file; qc_type = COVERAGE
    8. Parse the bed file
    9. Pass all data into the xlsx writer
    10. ???
    11. Profit
    """
    args = parse_sample_id_args()
    sample_id = args.s
    data_path = args.p
    snv_json = f'{data_path}/{sample_id}.hard-filtered.annotations.json.gz'
    sv_json = f'{data_path}/{sample_id}.sv.annotations.json.gz'
    cnv_json = f'{data_path}/{sample_id}.cnv.annotations.json.gz'
    cnv_seg = f'{data_path}/{sample_id}.seg.called.merged'
    cnv_vcf = f'{data_path}/{sample_id}.cnv.vcf.gz'
    tmb_csv = f'{data_path}/{sample_id}.tmb.metrics.csv'
    tmb_trace = f'{data_path}/{sample_id}.tmb.trace.tsv'
    summary_csv = f'{data_path}/Additional Files/{sample_id}.summary.csv'
    coverage_csv = f'{data_path}/{sample_id}.qc-coverage-region-1_coverage_metrics.csv'
    bed_file = f'{data_path}/{sample_id}.qc-coverage-region-1_read_cov_report.bed'
    snv_hits = parse_json(
        json_file=snv_json,
        var_type='SNV'
    )
    snp_id_hits = get_snp_id(snv_hits)
    sv_hits = parse_json(
        json_file=sv_json,
        var_type='SV'
    )
    seg_cnvs = parse_cnv_seg(
        cnv_seg=cnv_seg
    )
    json_cnvs = parse_cnv_json(
        cnv_json=cnv_json
    )
    cnv_hits = parse_cnvs(
        seg_cnvs=seg_cnvs,
        json_cnvs=json_cnvs
    )
    cnv_filter = parse_cnv_filtered(
        cnv_hits=cnv_hits
    )
    tmb = parse_qc_csv(
        csv_file=tmb_csv,
        qc_type='TMB'
    )
    summary = parse_qc_csv(
        csv_file=summary_csv,
        qc_type='SUMMARY'
    )
    coverage = parse_qc_csv(
        csv_file=coverage_csv,
        qc_type='COVERAGE'
    )
    bed = parse_bed(
        bed_file=bed_file
    )

    write_updated_cnv_vcf(cnv_filter, cnv_vcf)
    out_corrected_cnv_filtered(cnv_filter)
    snv_tmb_merge = add_tmb(tmb_trace, snv_hits)
    out_correct_cnv = out_corrected_cnv_filtered(cnv_filter)

    write_xlsx(
        data=(
            snv_hits,
            snp_id_hits,
            sv_hits,
            cnv_hits,
            out_correct_cnv,
            tmb,
            snv_tmb_merge,
            summary,
            coverage,
            bed
        ),
        sample_id=sample_id
    )
    write_xlsx_cnv_filtered(
        data=(
            out_correct_cnv,
        ),
        sample_id=sample_id
    )



if __name__ == '__main__':
    main()
