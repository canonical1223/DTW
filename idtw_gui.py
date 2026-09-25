#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IDTW Well Correlation — standalone Tkinter application (Python 3.10+).

Install: python -m pip install numpy pandas openpyxl
Run:     python idtw_gui.py
Checks:  python idtw_gui.py --self-test

All application and LAS parser code is in this file. No runtime downloads.
LAS MD may be M or FT (converted to metres); markers.xlsx MD must be metres.
Excel columns: Маркер, Скважина, UWI, MD, X, Y, Z. Coordinates may be blank.
UWI should be stored as text. Normal/consensus runs hold target markers out of
optimisation. The explicit tuning mode uses them to rank parameter settings;
its error is labelled as calibration, never independent validation.

Method: Fang et al. 2021, doi:10.1190/INT-2020-0172.1, eq.4.
Engineering choices are documented in correlate() and the GUI help.
Supports sequential pairwise multi-well correlation, ensemble alternatives,
manual picks and self-contained .idtw projects. It is not a global solver.
"""
import sys
if sys.version_info < (3, 10):
    raise SystemExit('Требуется Python 3.10 или новее.')
try:
    import numpy
    import pandas
    import openpyxl
except ImportError as dependency_error:
    raise SystemExit('Не найдена зависимость: ' + str(dependency_error) + '\n'
                     'Установите: python -m pip install numpy pandas openpyxl')


# BEGIN EMBEDDED LASFILE PARSER
# Source: https://github.com/bzlmnop/lasfile/blob/f6df63c709988abae045d5929cc661473fd73fce/src/lasfile/lasfile.py
# Upstream commit: f6df63c709988abae045d5929cc661473fd73fce
# Upstream lasfile.py SHA-256: a86bd7b04b6ca97ddc333b6b490ae56127dd7dd1f26b613eef7491d896f1436e
# Adaptations: embedded known_sections.json; optional apinum import.
# Parsing logic is otherwise unchanged. Application-specific validation follows below.
# This upstream component is distributed under the following MIT license.
# MIT License
# 
# Copyright (c) 2023 bzlmnop
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# A python module for reading LAS files

import json
import os
import re
import traceback
from io import StringIO
from numpy import genfromtxt
# from numpy import array
# from csv import reader
import pandas as pd
from pandas import DataFrame
from pandas import read_csv
import warnings
# Optional upstream feature: API-number formatting is not needed for correlation.
try:
    from apinum import APINumber
except ImportError:
    def APINumber(*args, **kwargs):
        raise ImportError("Install optional dependency 'apinum' to format API numbers.")



class LASFileError(Exception):
    """Raised when a LAS file experiences an error in the read, parse,
    or validata process"""
    pass


class LASFileCriticalError(LASFileError):
    """Raised when a LAS file experiences an error in the parse or
    validate process that is critical to the file's integrity"""
    pass


class LASVersionError(LASFileCriticalError):
    """Raised when a LAS file experiences an error in the version
    parsing and validation process"""
    pass


class UnknownVersionError(LASVersionError):
    """Raised when a version value extracted from an LAS file is not
    known"""
    pass


class VersionExtractionError(LASVersionError):
    """Raise when a version value cannot be extracted from an LAS"""
    pass


class LASFileOpenError(LASFileCriticalError):
    """Raised when a LAS file experiences an error in the open
    process"""
    pass


class LASFileReadError(LASFileCriticalError):
    """Raised when a LAS file experiences an error in the read
    process"""
    pass


class LASFileSplitError(LASFileCriticalError):
    """Raised when a LAS file experiences an error in the split
    process"""
    pass


class SectionTitleError(LASFileCriticalError):
    """Raised when a LAS file experiences an error in the section
    title parsing and validation process"""
    pass


class MissingRequiredSectionError(LASFileCriticalError):
    """Raised when a LAS file is missing a required section"""
    pass


class RequiredSectionParseError(LASFileCriticalError):
    """Raised when a required section in a LAS file fails to parse"""
    pass


class MissingCriticalMnemonicError(LASFileCriticalError):
    """Raised when a LAS file required section is missing a required
    mnemonic that is critical to properly parsing the file"""
    pass


class VersionValidationError(LASVersionError):
    """Raised when a LAS file experiences an error in the version
    validation process"""
    pass


class LASFileMinorError(LASFileError):
    """Raised when a LAS file experiences an error in the parse or
    validate process that is not critical to the file's integrity"""
    pass


class SectionLoadError(LASFileMinorError):
    """Raised when an error occurs when loading a non-required
    section of an LAS file"""
    pass


class SectionParseError(LASFileMinorError):
    """Raised when an error occurs when parsing a non-required
    section of an LAS file"""
    pass


class MissingMnemonicError(LASFileMinorError):
    """Raise when a LAS file required section is missing a required
    mnemonic that is not critical to properly parsing the file"""
    pass


# Set known versions
known_versions = ['1.2', '2.0', '3.0']
# Embedded upstream known_sections.json; no sidecar file is required.
known_secs = {'1.2': {'version': {'type': 'header', 'titles': ['v', 'V'], 'required': True},
         'well': {'type': 'header', 'titles': ['w', 'W'], 'required': True},
         'parameters': {'type': 'header', 'titles': ['p', 'P'], 'required': False},
         'curves': {'type': 'header', 'titles': ['c', 'C'], 'required': True},
         'data': {'type': 'data', 'titles': ['a', 'A'], 'required': True},
         'other': {'type': 'header', 'titles': ['o', 'O'], 'required': False}},
 '2.0': {'version': {'type': 'header', 'titles': ['v', 'V'], 'required': True},
         'well': {'type': 'header', 'titles': ['w', 'W'], 'required': True},
         'parameters': {'type': 'header', 'titles': ['p', 'P'], 'required': False},
         'curves': {'type': 'header', 'titles': ['c', 'C'], 'required': True},
         'data': {'type': 'data', 'titles': ['a', 'A'], 'required': True},
         'other': {'type': 'header', 'titles': ['o', 'O'], 'required': False}},
 '3.0': {'version': {'type': 'header', 'titles': ['version'], 'required': True},
         'well': {'type': 'header', 'titles': ['well'], 'required': True},
         'parameters': {'type': 'header',
                        'titles': ['parameter', 'log_parameter'],
                        'required': False},
         'curves': {'type': 'header',
                    'titles': ['curves', 'curve', 'curve_definition', 'log_definition'],
                    'required': True},
         'data': {'type': 'data', 'titles': ['log_data', 'ascii', 'data'], 'required': True},
         'core_parameters': {'type': 'header', 'titles': ['core_parameters'], 'required': False},
         'core_definition': {'type': 'header', 'titles': ['core_definition'], 'required': False},
         'core_data': {'type': 'data', 'titles': ['core_data'], 'required': False},
         'inclinometry_parameters': {'type': 'header',
                                     'titles': ['inclinometry_parameters'],
                                     'required': False},
         'inclinometry_definition': {'type': 'header',
                                     'titles': ['inclinometry_definition'],
                                     'required': False},
         'inclinometry_data': {'type': 'data', 'titles': ['inclinometry_data'], 'required': False},
         'drill_parameters': {'type': 'header', 'titles': ['drill_parameters'], 'required': False},
         'drill_definition': {'type': 'header', 'titles': ['drill_definition'], 'required': False},
         'drill_data': {'type': 'data', 'titles': ['drill_data'], 'required': False},
         'tops_parameters': {'type': 'header', 'titles': ['tops_parameters'], 'required': False},
         'tops_definition': {'type': 'header', 'titles': ['tops_definition'], 'required': False},
         'tops_data': {'type': 'data', 'titles': ['tops_data'], 'required': False},
         'test_parameters': {'type': 'header', 'titles': ['test_parameters'], 'required': False},
         'test_definition': {'type': 'header', 'titles': ['test_definition'], 'required': False},
         'test_data': {'type': 'data', 'titles': ['test_data'], 'required': False}}}

# Build a dictionary of required sections for each version from
# known_secs
required_sections = {}
for version, sections in known_secs.items():
    req_secs_list = [
        section_name for section_name, section in sections.items()
        if section['required']
    ]
    required_sections[version] = req_secs_list

header_section_names = []
data_section_names = []
for version, sections in known_secs.items():
    for section_name, sec_dict in sections.items():
        if (
            sec_dict['type'] == 'header' and
            section_name not in header_section_names
        ):
            header_section_names.append(section_name)
        elif (
            sec_dict['type'] == 'data' and
            section_name not in data_section_names
        ):
            data_section_names.append(section_name)


def get_version_num(data,
                    handle_common_errors=True,
                    accept_unknown_versions=False,
                    allow_non_numeric=False,
                    unknown_value=None):
    """
    Extracts and validates the version number from the given data.

    This function accepts either a string containing a version section,
    or a DataFrame containing a mnemonic column with a "VERS" value.
    It then tries to extract the version number and validate it.
    It handles several common errors, such as non-numeric versions, and
    allows the acceptance of unknown versions.

    Parameters:
    ----------
    data : str or pandas.DataFrame
        The input data to extract the version number from.
        If a string, it should contain a version section marked with '~V'.
        If a DataFrame, it should contain a column named 'mnemonic' with a
        "VERS" value.

    handle_common_errors : bool, optional
        Whether to handle common errors, such as whole number versions.
        (default is True)

    accept_unknown_versions : bool, optional
        Whether to accept and return unknown versions. (default is False)

    allow_non_numeric : bool, optional
        Whether to allow and return non-numeric versions. This only works
        if `accept_unknown_versions` is also True. (default is False)

    Returns:
    -------
    version_num : str
        The extracted version number.

    Raises:
    ------
    ValueError:
        If the input data is neither a string nor a DataFrame.

    Exception:
        If the version number could not be retrieved, or if it
        was not recognized and `accept_unknown_versions` is False.
    """
    # Parse input data based on its type
    if isinstance(data, str):
        try:
            # If the data is a string, extract the version section
            section_regex = re.compile(r'(~[V].+?)(?=~[VW]|$)', re.DOTALL)
            version_section = re.findall(section_regex, data)[0]
            # Parse the version section into a DataFrame
            df = parse_header_section(version_section)
        except Exception as e:
            raise LASVersionError(
                f"Could not extract and parse version section: {str(e)}"
            )
    elif isinstance(data, DataFrame):
        df = data
    else:
        raise ValueError("Input must be str or DataFrame.")

    # Try to extract version number
    try:
        version_num = df.loc[df['mnemonic'] == "VERS", "value"].values[0]
    except Exception as e:
        raise VersionExtractionError(f"Could not get version: {str(e)}")

    # Check if version number is known
    if version_num in known_versions:
        # Return version number if it is known
        return version_num

    # Handle common errors like conversion to float
    if handle_common_errors:
        try:
            float_version_num = str(float(version_num))
            if float_version_num in known_versions:
                return float_version_num
        except ValueError:
            pass

    # Accept unknown versions, verify if they are integers
    if accept_unknown_versions:
        try:
            float(version_num)
            return version_num
        except ValueError:
            if allow_non_numeric:
                return version_num

    # Raise error if no known version number was found even after
    # handling common errors and if unknown versions are not accepted
    raise UnknownVersionError(
        "Could not get version, version number not recognized."
    )


def get_version_section(data,
                        handle_common_errors=True,
                        accept_unknown_versions=False,
                        allow_non_numeric=False,
                        unknown_value=None):
    """
    Extracts the version section from the given raw data, parses it,
    validates it, and returns a loaded section object.

    This function performs several steps to process the version section
    from the raw data. It extracts the version section, parses it into
    a DataFrame, attempts to extract and validate a version number, and
    then tries to load this all into a LASSection object.

    Parameters:
    ----------
    data : str
        The raw data string to extract the version section from. The
        version section should be marked with '~V'.

    handle_common_errors : bool, optional
        Whether to handle common errors when extracting the
        version number. (default is True)

    accept_unknown_versions : bool, optional
        Whether to accept unknown versions when extracting the version
        number. (default is False)

    allow_non_numeric : bool, optional
        Whether to allow non-numeric versions when extracting the
        version number. (default is False)

    unknown_value : any, optional
        The value to use when an unknown version number is encountered.
        This only applies if `accept_unknown_versions`
        is also True. (default is None)

    Returns:
    -------
    loaded_section : LASSection
        The loaded version section.

    Raises:
    ------
    Exception:
        If parsing the version section fails, if extracting the
        version number fails, if validating the version section fails,
        or if loading the section into a LASSection object fails.
    """
    # Extract whole text of version section from raw data
    # Regex matches everything between '~V' and '~V' or '~W'
    section_regex = re.compile(r'(~[V].+?)(?=~[VW]|$)', re.DOTALL)
    section_list = re.findall(section_regex, data)
    # Take the first match, which should be the version section
    version_section = section_list[0]

    # Try to parse version section into a DataFrame
    try:
        df = parse_header_section(version_section)
    except Exception as e:
        raise RequiredSectionParseError(
            f"Failed to parse version section: {e}"
        )

    # Attempt to extract a version number from the parsed section
    try:
        version_num = get_version_num(
            df,
            handle_common_errors=handle_common_errors,
            accept_unknown_versions=accept_unknown_versions,
            allow_non_numeric=allow_non_numeric,
            unknown_value=unknown_value
        )
    except Exception as e:
        raise VersionExtractionError(f"Failed to extract version number: {e}")

    # Define default values for dlm_val and wrap
    dlm_val = None
    wrap = None

    # Attempt to validate the section
    if validate_version(df, version_num=version_num) == []:
        if version_num is not None:
            try:
                wrap_val = df.loc[df['mnemonic'] == "WRAP", "value"].values[0]
                if wrap_val.upper() == 'YES':
                    wrap = True
                elif wrap_val.upper() == 'NO':
                    wrap = False
                if "DLM" in df['mnemonic'].values:
                    dlm_val = df.loc[
                        df['mnemonic'] == "DLM", "value"].values[0]
            except Exception as e:
                wrap = None
                if version_num in ["1.2", "2.0"]:
                    raise MissingCriticalMnemonicError(
                        f"Could not get WRAP: {str(e)}"
                    )
        if version_num == "3.0":
            try:
                dlm_val = df.loc[df['mnemonic'] == "DLM", "value"].values[0]
            except Exception as e:
                dlm_val = None
                raise MissingCriticalMnemonicError(
                    f"Could not get DLM: {str(e)}"
                )

        # Attempt to load the section into a section object
        try:
            loaded_section = LASSection(
                'version',
                version_section,
                'header',
                version_num,
                delimiter=dlm_val,
                parse_on_init=False,
                validate_on_init=False,
                wrap=wrap
            )
            loaded_section.df = df
            loaded_section.validated = True
            loaded_section.version_num = version_num
            return loaded_section
        except Exception:
            raise LASVersionError("Couldn't load into section object.")
    else:
        raise VersionValidationError("Could not validate the version section.")


def parse_title_line(title_line,
                     version_num,
                     all_lowercase=True,
                     assocs=False):
    """
    Parses the title line based on the provided version number.

    This function handles parsing of title lines differently based
    on the given version number. Specifically, it handles versions
    1.2 and 2.0 differently from version 3.0.

    Parameters:
    ----------
    title_line : str
        The title line to be parsed.

    version_num : str
        The version number, which determines how the title line is
        parsed. Can be "1.2", "2.0", or "3.0".

    all_lowercase : bool, optional
        Whether to convert all characters in the title line to
        lowercase before parsing. (default is True)

    assocs : bool, optional
        Whether to consider associations in the parsing process.
        This is only relevant for version "3.0". (default is False)

    Returns:
    -------
    Parsed title line : varies
        The result of the title line parsing, which depends on the
        version number.

    Raises:
    ------
    Exception:
        If the version number is not one of the expected values
        ("1.2", "2.0", "3.0").
    """
    if version_num == "1.2" or version_num == "2.0":
        return parse_v2_title(title_line, all_lowercase=all_lowercase)
    elif version_num == "3.0":
        return parse_v3_title(
            title_line,
            all_lowercase=all_lowercase,
            assocs=assocs
        )


def parse_v2_title(title_line, all_lowercase=True):
    """
    Parses a version 2.0 title line.

    This function specifically handles the parsing of title lines for
    version 2.0. It checks if the line begins with '~', extracts the
    section title, and optionally converts it to lowercase.

    Parameters:
    ----------
    title_line : str
        The title line to be parsed. This should begin with '~'.

    all_lowercase : bool, optional
        Whether to convert the section title to lowercase.
        (default is True)

    Returns:
    -------
    section_title : str
        The parsed section title from the title line.

    Raises:
    ------
    Exception:
        If the title line does not begin with '~'.
    """
    # Strip leading and trailing whitespace
    title_line = title_line.strip()
    # Check if it actually is a title line
    if title_line.startswith('~'):
        title_line = title_line.strip('~')
        section_title = title_line[0]
        if all_lowercase:
            section_title = section_title.lower()
        return section_title
    else:
        raise SectionTitleError(
            "Cannot parse title line. Title lines must begin with '~'."
        )


def parse_v3_title(title_line, all_lowercase=True, assocs=False):
    """
    Parses a version 3.0 title line.

    This function specifically handles the parsing of title lines for
    version 3.0. It checks if the line begins with '~', extracts the
    section title and optionally an associated value, and converts them
    to lowercase if requested.

    Parameters:
    ----------
    title_line : str
        The title line to be parsed. This should begin with '~'.

    all_lowercase : bool, optional
        Whether to convert the section title and association to
        lowercase. (default is True)

    assocs : bool, optional
        Whether to consider associations in the parsing process.
        (default is False)

    Returns:
    -------
    section_title : str or tuple
        The parsed section title from the title line. If 'associations'
        is True and an association is present, a tuple of
        (section_title, association) is returned.

    Raises:
    ------
    Exception:
        If the title line does not begin with '~'.
    """
    assoc = None
    has_assoc = False
    # Strip leading and trailing whitespace
    title_line = title_line.strip()
    # Check if it actually is a title line
    if title_line.startswith('~'):
        # Strip leading '~'
        title_line = title_line.strip('~')
        # Check if there is an association
        if '|' in title_line:
            # If so, split the title line into the title and association
            has_assoc = True
            title_line, assoc = title_line.split('|')
            # If the association is not empty, strip leading and
            # trailing whitespace
            if assoc.strip().split(' ')[0] != '':
                assoc = assoc.strip().split(' ')[0]
            # Otherwise, set the association to None
            else:
                assoc = None
        # Extract the section title
        if title_line.split(' ')[0] != '':
            section_title = title_line.split(' ')[0]
        else:
            section_title = None
        # Convert to lowercase if requested
        if all_lowercase:
            if section_title is not None:
                section_title = section_title.lower()
            if assocs and has_assoc:
                if assoc is not None:
                    assoc = assoc.lower()
        # Return the section title and association if requested
        if assocs and has_assoc:
            return section_title, assoc
        # Otherwise, just return the section title
        else:
            return section_title
    else:
        # If the line does not begin with '~', raise an exception
        raise SectionTitleError(
            "Cannot parse a line that does not begin with '~' as a "
            "title line."
        )


def split_sections(data, version_num, known_secs=known_secs):
    """
    Splits the input data into sections based on the
    provided version number.

    This function handles the splitting of input data
    differently based on the given version number. It
    handles versions 1.2 and 2.0 differently from version
    3.0. For each version, it uses regular expressions to
    identify and extract the sections, and then stores these
    sections in a dictionary using appropriate keys.

    Parameters:
    ----------
    data : str
        The input data to be split into sections.

    version_num : str
        The version number, which determines how the
        data is split. Can be "1.2", "2.0", or "3.0".

    known_secs : dict, optional
        A dictionary mapping known section names to
        their details for each version.
        (default is known_secs)

    Returns:
    -------
    section_dict : dict
        A dictionary with the parsed sections. The keys
        depend on the version number and the content of
        the sections.

    Raises:
    ------
    Exception:
        If the version number is not one of the expected
        values ("1.2", "2.0", "3.0").
    """
    # Split the data into sections based on the version number
    if version_num == '1.2' or version_num == '2.0':
        # Use a regular expression to split the data into sections
        section_regex = re.compile(r'(~[VWPCOA].+?)(?=~[VWPCOA]|$)', re.DOTALL)
        sections = re.findall(section_regex, data)
        # Store the sections in a dictionary
        section_dict = {}
        for section in sections:
            # Get the title line/header of the section
            header_end = section.index('\n')
            header = section[:header_end].strip().upper()
            # Store the section in the dictionary based on the header
            if header.startswith('~W'):
                section_dict['well'] = section.strip()
            elif header.startswith('~V'):
                section_dict['version'] = section.strip()
            elif header.startswith('~C'):
                section_dict['curves'] = section.strip()
            elif header.startswith('~P'):
                section_dict['parameters'] = section.strip()
            elif header.startswith('~O'):
                section_dict['other'] = section.strip()
            elif header.startswith('~A'):
                section_dict['data'] = section.strip()
        # Return the dictionary of sections
        return section_dict
    elif version_num == '3.0':
        # Use a regular expression to split the data into sections
        section_regex = re.compile(r'(~[A-Za-z].+?)(?=~[A-Za-z]|$)', re.DOTALL)
        sections = re.findall(section_regex, data)
        # Store the sections in a dictionary
        section_dict = {}
        # Get the known sections for version 3.0
        known_secs = known_secs["3.0"]
        known_sec_names = known_secs.keys()
        for section in sections:
            # Get the title line/header of the section
            section_found = False
            try:
                title_line_end = section.index('\n')
            except Exception as e:
                raise LASFileCriticalError(
                    f"Could not find the end of "
                    f"the title line for a section: {e}"
                    )
            title_line = section[:title_line_end].strip()
            # Parse the title line to get the section title
            section_title = parse_title_line(
                title_line,
                "3.0",
                all_lowercase=True
            )
            # If the section title is a known section name, store it
            # using the known section name as the key
            if section_title in known_sec_names:
                section_dict[section_title] = section.strip()
                section_found = True
            # If the section title is not a known section name, check
            # if it is an alias for a known section name then store it
            # using the known section name as the key
            else:
                for known_sec_name, known_sec in known_secs.items():
                    if section_title in known_sec["titles"]:
                        section_dict[known_sec_name] = section.strip()
                        section_found = True
                if not section_found:
                    section_dict[section_title] = section.strip()
        # Return the dictionary of sections
        return section_dict
    else:
        # Raise an exception if the version number is not one of the
        # expected values
        raise UnknownVersionError(
            "Unknown version. Must be '1.2','2.0', or '3.0'"
        )


def parse_header_section(section_string, version_num='2.0', delimiter=None):
    """
    Parses the header section of a LAS file.

    This function parses the header section of a LAS file and returns
    a dictionary with the parsed information. The version number is
    used to determine how the header section is parsed. The delimiter
    is used to split the header section into lines. If no delimiter is
    provided, the header section is split into lines using the newline
    character.

    Parameters:
    ----------
    section_string : str
        The header section of a LAS file.

    version_num : str, optional
        The version number of the LAS file. (default is '2.0')

    delimiter : str, optional
        The delimiter used to split the header section into lines.
        (default is None)

    Returns:
    -------
    header_dict : dict
        A DataFrame with the parsed header information.
    """
    results = []
    lines = section_string.strip().split("\n")
    if version_num == '2.0' or version_num == '1.2':
        # Skip comment, title, and empty lines
        for line in lines:
            mnemonic = None
            units = None
            value = None
            descr = None
            if (
                line.strip().startswith("#") or
                line.strip().startswith("~") or
                line.strip().strip("\n") == ""
            ):
                continue
            # Try to parse the line
            try:
                # Remove whitespace from the line
                line = line.strip()
                # Get the mnemonic.
                # The mnemonic is everything before the first period,
                # stripped of whitespace.
                frst_prd = line.index('.')
                mnemonic = line[:frst_prd].strip()
                if version_num == '1.2' and mnemonic in [
                    'COMP',
                    'WELL',
                    'FLD',
                    'LOC',
                    'PROV',
                    'SRVC',
                    'DATE',
                    'UWI',
                    'API',
                ]:
                    # Get the units.
                    # The units are everything between the first period
                    # and the first space after the first period,
                    # stripped of whitespace.
                    line_aft_frst_prd = line[frst_prd+1:]
                    frst_spc_aft_frst_prd = line_aft_frst_prd.index(' ')
                    units = line_aft_frst_prd[:frst_spc_aft_frst_prd].strip()
                    # Get the value.
                    # The value is everything between the first space
                    # after the first period and the last colon,
                    # stripped of whitespace.
                    lst_col = line_aft_frst_prd.rindex(':')
                    descr = line_aft_frst_prd[
                        frst_spc_aft_frst_prd:lst_col
                    ].strip()
                    # Get the description.
                    # The description is everything after the last colon,
                    # stripped of whitespace.
                    value = line_aft_frst_prd[lst_col+1:].strip()
                    results.append(
                        {
                            "mnemonic": mnemonic,
                            "units": units if units != "" else None,
                            "value": value if value != "" else None,
                            "description": descr if descr != "" else None,
                            "errors": None
                        }
                    )
                else:
                    # Get the units.
                    # The units are everything between the first period
                    # and the first space after the first period,
                    # stripped of whitespace.
                    line_aft_frst_prd = line[frst_prd+1:]
                    frst_spc_aft_frst_prd = line_aft_frst_prd.index(' ')
                    units = line_aft_frst_prd[:frst_spc_aft_frst_prd].strip()
                    # Get the value.
                    # The value is everything between the first space
                    # after the first period and the last colon,
                    # stripped of whitespace.
                    lst_col = line_aft_frst_prd.rindex(':')
                    value = line_aft_frst_prd[
                        frst_spc_aft_frst_prd:lst_col
                    ].strip()
                    # Get the description.
                    # The description is everything after the last colon,
                    # stripped of whitespace.
                    descr = line_aft_frst_prd[lst_col+1:].strip()
                    results.append(
                        {
                            "mnemonic": mnemonic,
                            "units": units if units != "" else None,
                            "value": value if value != "" else None,
                            "description": descr if descr != "" else None,
                            "errors": None
                        }
                    )
            except Exception as e:
                results.append(
                        {
                            "mnemonic": mnemonic,
                            "units": units if units != "" else None,
                            "value": value if value != "" else None,
                            "description": descr if descr != "" else None,
                            "errors": LASFileError(
                                f"Error parsing header line '{line}'. {e}"
                            )
                        }
                    )
        return DataFrame(results)
    elif version_num == '3.0':
        for line in lines:
            mnemonic = None
            units = None
            value = None
            descr = None
            format = None
            assocs = None
            # Skip comment, title, and empty lines
            if (
                line.strip().startswith("#") or
                line.strip().startswith("~") or
                line.strip().strip("\n") == ""
            ):
                continue
            # Try to parse the line
            try:
                # Remove whitespace from the line
                line = line.strip()
                # Get the mnemonic.
                # The mnemonic is everything before the
                # first period, stripped of whitespace
                frst_prd = line.index('.')
                mnemonic = line[:frst_prd].strip()
                # Get the units.
                # The units are everything between the first period and the
                # first space after the first period, stripped of whitespace.
                line_aft_frst_prd = line[frst_prd+1:]
                frst_spc_aft_frst_prd = line_aft_frst_prd.index(' ')
                units = line_aft_frst_prd[:frst_spc_aft_frst_prd].strip()
                # Get the value.
                # The value is everything between the first space after the
                # first period and the last colon, stripped of whitespace.
                lst_col = line_aft_frst_prd.rindex(':')
                value = (
                    line_aft_frst_prd[frst_spc_aft_frst_prd:lst_col].strip()
                )
                # Get the description, format, and associations.
                # The description is everything after the last colon and
                # before the first brace or pipe, if no braces are present,
                # stripped of whitespace.
                line_aft_lst_col = line_aft_frst_prd[lst_col+1:]
                if '{' in line_aft_lst_col and '}' in line_aft_lst_col:
                    frst_brc_aft_lst_col = line_aft_lst_col.index('{')
                    descr = line_aft_lst_col[:frst_brc_aft_lst_col]
                    line_aft_frst_brc = (
                        line_aft_lst_col[frst_brc_aft_lst_col+1:]
                    )
                    clsng_brc = line_aft_frst_brc.rindex('}')
                    format = line_aft_frst_brc[:clsng_brc].strip()
                    line_aft_clsng_brc = line_aft_frst_brc[clsng_brc+1:]
                    if '|' in line_aft_clsng_brc:
                        bar = line_aft_clsng_brc.index('|')
                        assocs = line_aft_clsng_brc[bar+1:].strip()
                elif '|' in line_aft_lst_col:
                    bar = line_aft_lst_col.index('|')
                    assocs = line_aft_lst_col[bar+1:].strip()
                else:
                    descr = line_aft_frst_prd[lst_col+1:].strip()
                # Add the parsed values to the results list
                results.append(
                    {
                        "mnemonic": mnemonic,
                        "units": units if units != "" else None,
                        "value": value if value != "" else None,
                        "description": descr if descr != "" else None,
                        "format": format if format != "" else None,
                        "associations": assocs if assocs != "" else None
                    }
                )
            except Exception as e:
                results.append(
                        {
                            "mnemonic": mnemonic,
                            "units": units if units != "" else None,
                            "value": value if value != "" else None,
                            "description": descr if descr != "" else None,
                            "errors": LASFileError(
                                f"Error parsing header line '{line}'. {e}"
                            )
                        }
                    )
        # Return the results as a DataFrame
        return DataFrame(results)


def parse_data_section(
        raw_data,
        version_num,
        wrap,
        delimiter=None,
        curve_names=None
):
    """
    Parses the data section of the input raw data based on the provided
    version number.

    This function handles the parsing of the data section from the
    input raw data. It removes lines beginning with '#' or '~', and
    then loads the data into a LASData object.

    Parameters:
    ----------
    raw_data : str
        The raw data to be parsed.

    version_num : str
        The version number, which is used when loading the data into a
        LASData object.

    delimiter : str, optional
        The delimiter character used to separate values in the data. If
        not provided, the default delimiter for the LASData class is
        used.

    Returns:
    -------
    loaded_data : LASData
        The loaded data as a LASData object.
    """
    filtered_data = re.sub(r'^[#~].*\n', '', raw_data, flags=re.MULTILINE)
    loaded_data = LASData(
        filtered_data,
        version_num,
        wrap=wrap,
        delimiter=delimiter,
        curve_names=curve_names)
    return loaded_data


def validate_version(df, version_num=None):
    """
    Validates the version of a dataframe.

    This function checks if the dataframe contains all required
    mnemonics for the given version number and verifies the values of
    these mnemonics. The specific checks differ based on the version
    number.

    Parameters:
    ----------
    df : DataFrame
        The dataframe to be validated. This dataframe should contain a
        'mnemonic' column and a 'value' column.

    version_num : str, optional
        The version number of the dataframe. Must be either "1.2",
        "2.0", or "3.0". If not provided, the function will use the
        value in the dataframe.

    Returns:
    -------
    bool:
        Returns True if the dataframe is valid for the given version
        number.

    Raises:
    ------
    Exception:
        If any of the required mnemonics are missing or have invalid
        values, or if the version number is not one of the expected
        values ("1.2", "2.0", "3.0").
    """
    validate_errors = []
    if version_num in ["1.2", "2.0"]:
        req_mnemonics = ["VERS", "WRAP"]
    elif version_num == "3.0":
        req_mnemonics = ["VERS", "WRAP", "DLM"]
    else:
        validate_errors.append(
            UnknownVersionError(
                "Unknown version. Must be '1.2','2.0', or '3.0'"
            )
        )
        return validate_errors

    # Test if all required mnemonics are present
    if not all(mnemonic in df.mnemonic.values for mnemonic in req_mnemonics):
        # Try and fix the missing mnemonic error by adjusting mnemonic
        # case to upper.
        if not all(
            mnemonic in [val.upper() for val in df.mnemonic.values]
            for mnemonic in req_mnemonics
        ):
            # Make a list of the missing mnemonics
            missing_mnemonics = [
                mnemonic for mnemonic in req_mnemonics
                if mnemonic not in df.mnemonic.values
            ]
            validate_errors.append(
                MissingCriticalMnemonicError(
                    f"Missing required version section mnemonics: "
                    f"{missing_mnemonics}"
                )
            )
            return validate_errors
        else:
            # Auto repair the mnemonic case
            df['mnemonic'] = df['mnemonic'].str.upper()

    # Set empty wrap value
    wrap = None

    try:
        wrap = df.loc[df['mnemonic'] == "WRAP", "value"].values[0]
    except Exception as e:
        if version_num in ["1.2", "2.0"]:
            validate_errors.append(
                LASVersionError(f"Couldnt get WRAP value: {str(e)}")
            )
            return validate_errors
        else:
            pass

    if version_num in ["1.2", "2.0"]:
        if wrap is not None and wrap.upper() not in ["YES", "NO"]:
            validate_errors.append(
                LASVersionError(
                    "Wrap value for versions 1.2 and 2.0 must be 'YES' "
                    "or 'NO'."
                )
            )
            return validate_errors
    elif version_num == "3.0":
        try:
            dlm = df.loc[df['mnemonic'] == "DLM", "value"].values[0]
            if "wrap" in locals():
                if wrap is not None and wrap.upper() != "NO":
                    validate_errors.append(
                        LASVersionError(
                            "Invalid wrap value. Must be 'NO' for version 3.0"
                        )
                    )
                    return validate_errors
            if dlm.upper() not in ["SPACE", "COMMA", "TAB", None, '']:
                validate_errors.append(
                    LASVersionError(
                        "Invalid delimiter value for version 3.0, should be "
                        "'SPACE', 'COMMA', or 'TAB'"
                    )
                )
                return validate_errors
        except Exception as e:
            validate_errors.append(
                LASVersionError(f"Couldnt get DLM value: {str(e)}")
            )
            return validate_errors
    return validate_errors


def validate_v2_well(df):
    """
    Validates the well section of a dataframe for version 2 LAS files.

    This function checks if the dataframe contains all the required
    mnemonics specific to the well section of a version 2 LAS file.

    Parameters:
    ----------
    df : DataFrame
        The dataframe to be validated. This dataframe should contain a
        'mnemonic' column.

    Returns:
    -------
    bool:
        Returns True if the dataframe is valid for the well section of
        a version 2 LAS file.

    Raises:
    ------
    Exception:
        If any of the required mnemonics are missing.
    """
    validate_errors = []
    # Set of required mnemonics for version 2.0 well sections
    req_mnemonics = [
        "STRT",
        "STOP",
        "STEP",
        "NULL",
        "COMP",
        "WELL",
        "FLD",
        "LOC",
        "SRVC",
        "DATE",
    ]
    # Instantiate an empty list to store missing mnemonics
    missing_mnemonics = []
    # Check if all required mnemonics are present
    if all(mnemonic not in df.mnemonic.values for mnemonic in req_mnemonics):
        # Make a list of which mnemonics are missing
        for mnemonic in req_mnemonics:
            if mnemonic not in df.mnemonic.values:
                missing_mnemonics.append(mnemonic)
    # Check that either PROV or CNTY, STAT, CTRY required mnemonics
    # are present
    if (
        "PROV" not in df.mnemonic.values and
        all(
            mnemonic not in df.mnemonic.values
            for mnemonic in ["CNTY", "STAT", "CTRY"]
        )
    ):
        # Make a list of which mnemonics are missing
        for mnemonic in ["CNTY", "STAT", "CTRY"]:
            if mnemonic not in df.mnemonic.values:
                missing_mnemonics.append(mnemonic)
    if (
            "API" not in df.mnemonic.values and
            "UWI" not in df.mnemonic.values
    ):
        missing_mnemonics.append("API")
        missing_mnemonics.append("UWI")
    if missing_mnemonics != []:
        validate_errors.append(
            MissingMnemonicError(
                f"Missing required mnemonics: {missing_mnemonics}"
            )
        )
    # Test if there is an errors column in the dataframe
    if "errors" in df.columns:
        # If the value in the mnemonic column is in the req_mnemonics
        # list, and the value in the errors column is not None, append
        # a the error to the validate_errors list
        for index, row in df.iterrows():
            # Use pd.isna() to check for both None and NaN
            errors_val = row["errors"]
            is_error = errors_val is not None and not (
                isinstance(errors_val, float) and pd.isna(errors_val)
            )
            if row["mnemonic"] in req_mnemonics and is_error:
                validate_errors.append(
                    LASFileCriticalError(
                        f"Error parsing required header line "
                        f"'{index}', {errors_val}"
                    )
                )
            elif is_error:
                validate_errors.append(
                    LASFileMinorError(
                        f"Error parsing header line '{index}', "
                        f"{errors_val}"
                    )
                )
    return validate_errors


def validate_v3_well(df):
    """
    Validates the well section of a DataFrame for
    version 3 LAS files.

    This function checks if the DataFrame contains all the required
    mnemonics specific to the well section of a version 3 LAS file. It
    also validates geographic coordinates and country code.

    Parameters:
    ----------
    df : DataFrame
        The DataFrame to be validated. This DataFrame should contain a
        'mnemonic' column.

    Returns:
    -------
    bool:
        Returns True if the DataFrame is valid for the well section of
        a version 3 LAS file.

    Raises:
    ------
    Exception:
        If any of the required mnemonics are missing, or if geographic
        coordinates or country code are not properly specified.
    """
    validate_errors = []
    # Set of required mnemonics for version 3.0 well sections
    req_mnemonics = [
        "STRT",
        "STOP",
        "STEP",
        "NULL",
        "COMP",
        "WELL",
        "FLD",
        "LOC",
        "SRVC",
        "CTRY",
        "DATE",
    ]
    # Set of valid country codes for the CTRY mnemonic
    valid_country_codes = ["US", "CA"]
    # Instantiate an empty list to store missing mnemonics
    missing_mnemonics = []
    # Check if all required mnemonics are present
    if all(
        mnemonic not in df.mnemonic.values
        for mnemonic in req_mnemonics
    ):
        # Make a list of which mnemonics are missing
        for mnemonic in req_mnemonics:
            if mnemonic not in df.mnemonic.values:
                missing_mnemonics.append(mnemonic)
    # Check that either LATI, LONG, GDAT or X, Y, GDAT, HZCS are present
    if (
        all(
            mnemonic not in df.mnemonic.values
            for mnemonic in ["LATI", "LONG", "GDAT"]
        )
        or
        all(
            mnemonic not in df.mnemonic.values
            for mnemonic in ["X", "Y", "GDAT", "HZCS"]
        )
    ):
        if (
            any(
                mnemonic in df.mnemonic.values
                for mnemonic in ["LATI", "LONG", "GDAT"]
            )
        ):
            # Make a list of which mnemonics are missing
            for mnemonic in ["LATI", "LONG", "GDAT"]:
                if mnemonic not in df.mnemonic.values:
                    missing_mnemonics.append(mnemonic)
        elif (
            any(
                mnemonic in df.mnemonic.values
                for mnemonic in ["X", "Y", "GDAT", "HZCS"]
            )
        ):
            # Make a list of which mnemonics are missing
            for mnemonic in ["X", "Y", "GDAT", "HZCS"]:
                if mnemonic not in df.mnemonic.values:
                    missing_mnemonics.append(mnemonic)
    # Check that CTRY is present and a valid valued and the required
    # mnemonics for the country code are present
    if "CTRY" in df.mnemonic.values:
        country_code = df.loc[
            df["mnemonic"] == "CTRY", 'value'
        ].values[0]
        if country_code is not None:
            country_code = country_code.upper()
            if country_code in valid_country_codes:
                if country_code == "CA":
                    if all(
                        mnemonic not in df.mnemonic.values
                        for mnemonic in ["PROV", "UWI", "LIC"]
                    ):
                        # Make a list of which mnemonics are missing
                        for mnemonic in ["PROV", "UWI", "LIC"]:
                            if mnemonic not in df.mnemonic.values:
                                missing_mnemonics.append(mnemonic)
                elif country_code == "US":
                    if all(
                        mnemonic not in df.mnemonic.values
                        for mnemonic in ["STAT", "CNTY", "API"]
                    ):
                        # Make a list of which mnemonics are missing
                        for mnemonic in ["STAT", "CNTY", "API"]:
                            if mnemonic not in df.mnemonic.values:
                                missing_mnemonics.append(mnemonic)
                elif (
                        country_code is None or
                        country_code == ''
                ):
                    pass
                else:
                    validate_errors.append(
                        LASFileMinorError(
                            "Value for country code mnemonic is invalid: "
                            f"{country_code}. Must be a valid internet "
                            "country code."
                        )
                    )
        else:
            validate_errors.append(
                LASFileMinorError(
                    "Value for country code is missing."
                )
            )
    if missing_mnemonics != []:
        validate_errors.append(
            MissingMnemonicError(
                f"Missing required mnemonics: {missing_mnemonics}"
            )
        )
    # Test if there is an errors column in the dataframe
    if "errors" in df.columns:
        # If the value in the mnemonic column is in the req_mnemonics
        # list, and the value in the errors column is not None, append
        # a the error to the validate_errors list
        for index, row in df.iterrows():
            # Use pd.isna() to check for both None and NaN
            errors_val = row["errors"]
            is_error = errors_val is not None and not (
                isinstance(errors_val, float) and pd.isna(errors_val)
            )
            if row["mnemonic"] in req_mnemonics and is_error:
                validate_errors.append(
                    LASFileCriticalError(
                        f"Error parsing required header line "
                        f"'{index}', {errors_val}"
                    )
                )
            elif is_error:
                validate_errors.append(
                    LASFileMinorError(
                        f"Error parsing header line '{index}', "
                        f"{errors_val}"
                    )
                )
    return validate_errors


def validate_well(df, version_num):
    """
    Validates the well section of a DataFrame for specified LAS file
    versions.

    This function calls either validate_v2_well or validate_v3_well
    depending on the version number provided. It verifies the DataFrame
    for the required structure according to the LAS version.

    Parameters:
    ----------
    df : DataFrame
        The DataFrame to be validated. This DataFrame should contain a
        'mnemonic' column.

    version_num : str
        A string representing the LAS version. Valid values are "1.2",
        "2.0", and "3.0".

    Returns:
    -------
    bool:
        Returns True if the DataFrame is valid for the well section of
        the given LAS version.

    Raises:
    ------
    Exception:
        If validation fails in the respective validate function.
    """
    validate_errors = []
    if version_num == "1.2" or version_num == "2.0":
        try:
            errors = validate_v2_well(df)
            if errors is not None:
                for error in errors:
                    validate_errors.append(error)
        except Exception as e:
            validate_errors.append(
                LASFileCriticalError(
                    f"Error validating well section: {e}"
                )
            )
    if version_num == "3.0":
        try:
            errors = validate_v3_well(df)
            if errors is not None:
                for error in errors:
                    validate_errors.append(error)
        except Exception as e:
            validate_errors.append(
                LASFileCriticalError(
                    f"Error validating well section: {e}"
                )
            )
    return validate_errors


def validate_curves(df, version_num):
    """
    Validates the curves section of a specified LAS file.

    This function checks if the DataFrame contains all the required
    mnemonics specific to the curves section of a LAS file. It also
    checks that the mnemonics are in the correct order.

    Parameters:
    ----------
    df : DataFrame
        The DataFrame of the parsed curves section of a LAS file to be
        validated. This DataFrame should contain a 'mnemonic' column and
        an error column.

    version_num : str
        A string representing the LAS version. Valid values are "1.2",
        "2.0", and "3.0".

    Returns:
    -------
    bool:
        Returns True if the DataFrame is valid for the curves section
        of the given LAS version.
    """
    validate_errors = []
    # If ther is an errors column, check that there are no errors
    if "errors" in df.columns:
        if not df.errors.isnull().all():
            # If there are errors, append them to the validate_errors
            # list
            for index, row in df.iterrows():
                if row["errors"] is not None:
                    validate_errors.append(
                        LASFileCriticalError(
                            f"Error parsing curve line '{index}', "
                            f"{row['errors']}"
                        )
                    )
    return validate_errors


def unwrap_las_data(num_values, las_data):
    """
    Takes wrapped data, where a single depth value/row is split across
    multiple lines, and unwraps it into a single line per depth value.

    Parameters:
    ----------
    num_values : int
        The number of curves/values that should be present in each row
        of the unwrapped data.

    las_data : str
        The raw wrapped LAS data to be unwrapped.

    Returns:
    -------
    unwrapped_data : str
        The unwrapped LAS data.
    """
    # split the data into individual lines
    lines = las_data.split('\n')
    unwrapped_data = []
    current_row = []
    for line in lines:
        numbers = re.findall(r'\S+', line)
        # append columns to the current row
        current_row.extend(numbers)
        # if row has the expected number of values
        if len(current_row) == num_values:
            # append the row
            unwrapped_data.append(' '.join(current_row) + '\n')
            # start a new row
            current_row = []
    # join all unwrapped rows into a single string
    return ''.join(unwrapped_data)


class LASData():
    """
    Class for storing and handling Log ASCII
    Standard (LAS) data.

    The class provides an interface to load, handle and validate LAS
    data. The LAS format is used to store well log data in the oil and
    gas industry.

    Attributes:
    ----------
    raw_data : str
        The raw LAS data as a string.

    version_num : str
        The version of the LAS file.

    wrap : bool

    delimiter : str, optional
        The delimiter used in the LAS data, such as SPACE, COMMA, or
        TAB. Defaults to None, indicating a space delimiter.

    invalid_raise : bool, optional
        Whether to raise an exception when invalid data is encountered.
        Defaults to False.

    unrecognized_delimiters : bool, optional
        If True, any unrecognized delimiters are replaced with the
        default delimiter.
        Defaults to True.

    default_delimiter : str, optional
        The default delimiter to use if an unrecognized delimiter is
        found. Defaults to a space.

    delimiter_error : str, optional
        Stores an error message if an unrecognized delimiter is found.

    data : numpy.ndarray
        The parsed LAS data, stored as a numpy array.

    df : pandas.DataFrame
        The parsed LAS data, stored as a pandas DataFrame.

    read_errors : list, optional
        Stores any read errors encountered when loading the data.

    Methods:
    -------
    __init__(self, raw_data, version_num, delimiter=None,
    invalid_raise=False, unrecognized_delimiters=True,
    default_delimiter=' ')
        Initializes the LASData object by parsing the provided
        raw LAS data.
    """
    def __init__(
        self,
        raw_data,
        version_num,
        wrap=False,
        delimiter=None,
        invalid_raise=False,
        unrecognized_delimiters=True,
        default_delimiter=' ',
        curve_names=None
    ):
        # Initialize attributes
        self.raw_data = raw_data
        self.version_num = version_num
        self.wrap = wrap
        self.delimiter = delimiter
        # Initialize the delimiter dictionary
        delim_dict = {
            'SPACE': ' ',
            'COMMA': ',',
            'TAB': '\t'
        }
        delim = None
        # Get the delimiter value from the dictionary
        if self.delimiter is not None and self.delimiter in delim_dict.keys():
            delim = delim_dict[self.delimiter]
        # If the delimiter is a space, comma, tab, or None, use it
        elif self.delimiter in delim_dict.values() or self.delimiter is None:
            delim = self.delimiter
        elif unrecognized_delimiters:
            # If unrecognized delimiters are allowed, use the default
            # delimiter if it is in the dictionary
            if default_delimiter in delim_dict.values():
                delim = default_delimiter
            else:
                # Otherwise, raise an error
                self.delimiter_error = LASFileCriticalError(
                    f"Unrecognized delimiter: '{self.delimiter}', and default "
                    f"delimiter '{default_delimiter} unable to load!"
                )
        else:
            # If unrecognized delimiters are not allowed, and the
            # input delimiter is not in the dictionary, raise an error
            self.delimiter_error = (
                LASFileCriticalError(
                    f"Unrecognized delimiter: '{self.delimiter}', "
                    "unable to load!"
                )
            )
            return
        if wrap:
            if curve_names is not None:
                num_values = len(curve_names)
                self.unwrapped_data = unwrap_las_data(
                    num_values,
                    self.raw_data
                )
                with StringIO(self.unwrapped_data) as f:
                    # Catch any warnings that occur when reading the data
                    with warnings.catch_warnings(record=True) as w:
                        warnings.simplefilter("always")
                        # Use numpy's genfromtxt to read the data into a
                        # numpy array
                        self.data = genfromtxt(
                            f,
                            delimiter=delim,
                            invalid_raise=invalid_raise
                        )
                        # Convert the numpy array to a pandas DataFrame
                        self.df = DataFrame(self.data)
                        for warn in w:
                            if issubclass(warn.category, UserWarning):
                                # Store any read errors
                                self.read_errors = [
                                    err
                                    for err
                                    in str(warn.message).split("\n")
                                ]
                                return
            else:
                self.read_errors = [
                    LASFileCriticalError(
                        "Curve names must be provided when wrap is True"
                    )
                ]
        # If the data is not wrapped...
        else:
            # Create a file-like object from the string
            with StringIO(self.raw_data) as f:
                # Catch any warnings that occur when reading the data
                with warnings.catch_warnings(record=True) as w:
                    warnings.simplefilter("always")
                    # Use numpy's genfromtxt to read the data into a
                    # numpy array
                    if delim == ' ':
                        # self.data = genfromtxt(
                        #     f,
                        #     invalid_raise=invalid_raise
                        # )
                        self.df = read_csv(
                            f,
                            sep=r"\s+",
                            header=None
                        )
                    elif delim == ',':
                        # csv_data = reader(f, delimiter=delim)
                        # csv_data = list(csv_data)
                        # self.data = array(csv_data)
                        self.df = read_csv(
                            f,
                            delimiter=delim,
                            header=None
                        )
                    elif delim == '\t':
                        # self.data = genfromtxt(
                        #     f,
                        #     delimiter=delim,
                        #     invalid_raise=invalid_raise
                        # )
                        self.df = read_csv(
                            f,
                            delimiter=delim,
                            header=None
                        )
                    else:
                        self.df = read_csv(
                            f,
                            sep=r"\s+",
                            header=None
                        )
                    # Convert the numpy array to a pandas DataFrame
                    # self.df = DataFrame(self.data)
                    for warn in w:
                        if issubclass(warn.category, UserWarning):
                            # Store any read errors
                            self.read_errors = [
                                err
                                for err
                                in str(warn.message).split("\n")
                            ]
                            return


class LASSection():
    """
    Class representing a section of Log ASCII
    Standard (LAS) data.

    This class provides methods and properties
    for parsing, validating and handling a LAS section.

    Attributes:
    ----------
    name : str
        The name of the section.

    raw_data : str
        The raw data of the section as a string.

    type : str
        The type of the section, either 'header', 'data', or 'other'.

    version_num : str
        The version number of the LAS file the section belongs to.

    assoc : str, optional
        The association of the section.
        Defaults to None.

    delimiter : str, optional
        The delimiter used in the section.
        Defaults to None.

    parsed_section : obj, optional
        The parsed section. This is populated when the section is
        parsed.

    parsed : bool, optional
        Indicates whether the section has been parsed.
        Defaults to False.

    validated : bool, optional
        Indicates whether the section has been validated.
        Defaults to False.

    df : pandas.DataFrame
        The DataFrame representing the parsed data of the section.

    Methods:
    -------
    __init__(self, name, raw_data, section_type, version_num,
    assoc=None, delimiter=None, parse_on_init=True,
    validate_on_init=True)
        Initializes the LASSection object, parses the raw data,
        validates the parsed section if parse_on_init and
        validate_on_init are set to True.

    __repr__(self)
        Returns a string that represents the LASSection object in a way
        that can be used to recreate the object.

    __str__(self)
        Returns a user-friendly string representation of the LASSection
        object.

    parse(self)
        Parses the raw data into a usable format (not implemented in
        this code snippet).

    validate(self)
        Validates the parsed section (not implemented in this code
        snippet).

    """
    def __init__(
        self,
        name,
        raw_data,
        section_type,
        version_num,
        wrap,
        delimiter=None,
        assoc=None,
        curve_names=None,
        parse_on_init=True,
        validate_on_init=True
    ):
        # initialize attributes
        self.name = name
        self.raw_data = raw_data
        self.type = section_type
        self.version_num = version_num
        self.association = assoc
        self.delimiter = delimiter
        self.validated = False
        self.wrap = wrap
        self.curve_names = curve_names
        # parse section
        if parse_on_init:
            self.parse_errors = []
            self.parse_tbs = []
            try:
                self.parse()
            except Exception as e:
                self.parse_errors.append(
                    LASFileError(
                        f"Couldn't parse section {self.name}: {str(e)}"
                    )
                )
                self.parse_tbs.append(traceback.format_exc())
                return
            if self.parse_errors == []:
                del self.parse_errors
            if self.parse_tbs == []:
                del self.parse_tbs
        # validate section
        if (
            parse_on_init and validate_on_init
        ):
            self.validate_errors = []
            self.validate_tbs = []
            # If there are no parse errors, proceed normally with validation
            if not hasattr(self, 'parse_errors') or self.parse_errors == []:
                try:
                    self.validate()
                    if self.validate_errors != []:
                        self.validated = False
                    else:
                        self.validated = True
                except Exception as e:
                    self.validated = False
                    self.validate_errors.append(e)
                    self.validate_tbs.append(traceback.format_exc())
            # If there are parse errors...
            else:
                # If there are critical parse errors, return a critical
                # validation error
                if any(
                    isinstance(error, LASFileCriticalError)
                    for error in getattr(self, 'parse_errors')
                ):
                    self.validate_errors.append(
                        LASFileCriticalError(
                            "Couldn't validate section due to critical parse "
                            "errors."
                        )
                    )
                # If there are no critical parse errors...
                else:
                    # attempt validation
                    try:
                        self.validate()
                        if self.validate_errors != []:
                            self.validated = False
                        else:
                            self.validated = True
                    except Exception as e:
                        self.validate_errors.append(e)
                        self.validate_tbs.append(traceback.format_exc())
            if self.validate_errors == []:
                del self.validate_errors
            if self.validate_tbs == []:
                del self.validate_tbs

    def parse(self):
        """
        Parses the raw data of the section into a usable format.

        This method parses the raw data of the section into a usable
        format. The specific parsing steps differ based on the version
        number and the type of the section.

        Parameters:
        ----------
        None

        Returns:
        -------
        None
        """
        if '\n' in self.raw_data:
            # parse title line
            title_line_end = self.raw_data.index('\n')
            title_line = self.raw_data[:title_line_end].strip()
            if '|' in title_line:
                result = parse_title_line(
                    title_line,
                    version_num=self.version_num,
                    all_lowercase=True,
                    assocs=True
                )
                if type(result) is not str and result is not None:
                    self.association = result[1]
                else:
                    self.association = None
            else:
                self.association = None
            # if the section is a header section, parse it as such
            if self.type.lower() == 'header':
                try:
                    self.parsed_section = parse_header_section(
                        self.raw_data,
                        version_num=self.version_num
                    )
                    self.df = self.parsed_section
                    # Test if there are any errors in the parsed section
                    # by check if the only value in the errors column is
                    # None
                    if 'errors' in self.df.columns:
                        if self.df['errors'].unique().tolist() == [None]:
                            # If there are no errors remove the errors column
                            self.df = self.df.drop(columns=['errors'])
                        else:
                            for error in self.df['errors']:
                                if error is not None:
                                    self.parse_errors.append(
                                        error
                                    )
                except Exception as e:
                    # if it is a required section, return a critical error
                    # otherwise return a minor error
                    if (
                        self.name.lower() in
                        required_sections[self.version_num]
                    ):
                        self.parse_errors.append(
                            RequiredSectionParseError(
                                f"Couldn't parse '{self.name}' data: {str(e)}"
                            )
                        )
                        self.parse_tbs.append(traceback.format_exc())
                        return
                    else:
                        self.parse_errors.append(
                            SectionParseError(
                                f"Couldn't parse '{self.name}' data: {str(e)}"
                            )
                        )
                        self.parse_tbs.append(traceback.format_exc())
                        return
            # if the section is a data section, parse it as such
            elif self.type.lower() == 'data':
                try:
                    # print(f"section name: {name}")
                    self.parsed_section = parse_data_section(
                        self.raw_data,
                        version_num=self.version_num,
                        wrap=self.wrap,
                        delimiter=self.delimiter,
                        curve_names=self.curve_names
                    )
                    self.df = self.parsed_section.df
                    # Test if there are any errors in the parsed section
                    # by check if the only value in the errors column is
                    # None
                    if 'errors' in self.df.columns:
                        if (
                            getattr(
                                self, 'df'
                            )['errors'].unique().tolist() == [None]
                        ):
                            # If there are no errors remove the errors column
                            self.df = self.df.drop(columns=['errors'])
                        else:
                            for error in self.df['errors']:
                                self.parse_errors.append(
                                    error
                                )
                except Exception as e:
                    # if it is a required section, return a critical error
                    # otherwise return a minor error
                    if (
                        self.name.lower() in
                        required_sections[self.version_num]
                    ):
                        self.parse_errors.append(
                            RequiredSectionParseError(
                                f"Couldn't parse '{self.name}' data: {str(e)}"
                            )
                        )
                        self.parse_tbs.append(traceback.format_exc())
                        return
                    else:
                        self.parse_errors.append(
                            SectionParseError(
                                f"Couldn't parse '{self.name}' data: {str(e)}"
                            )
                        )
                        self.parse_tbs.append(traceback.format_exc())
                        return
            # parse other sections
            else:
                try:
                    self.parsed_section = parse_header_section(
                        self.raw_data,
                        version_num=self.version_num
                        )
                    self.df = self.parsed_section
                    # Test if there are any errors in the parsed section
                    if 'errors' in self.df.columns:
                        if self.df['errors'].unique().tolist() == [None]:
                            # If there are no errors remove the errors column
                            self.df = self.df.drop(columns=['errors'])
                            self.type = 'header'
                        else:
                            raise Exception('Not a header section')
                except Exception:
                    try:
                        self.parsed_section = parse_data_section(
                            self.raw_data,
                            version_num=self.version_num,
                            wrap=self.wrap,
                            delimiter=self.delimiter
                        )
                        self.df = self.parsed_section.df
                        self.type = 'data'
                    except Exception as e:
                        # Test if it is a required section or not and
                        # return the appropriate error
                        if (
                            self.name.lower() in
                            required_sections[self.version_num]
                        ):
                            self.parse_errors.append(
                                RequiredSectionParseError(
                                    f"Couldn't parse '{self.name}' data: "
                                    f"{str(e)}"
                                )
                            )
                            self.parse_tbs.append(traceback.format_exc())
                        else:
                            self.parse_errors.append(
                                SectionParseError(
                                    f"Couldn't parse '{self.name}' data: "
                                    f"{str(e)}"
                                )
                            )
                            self.parse_tbs.append(traceback.format_exc())
                        return
        # if raw data is only one line or less
        else:
            # Test if it is a critical required section or not and
            # return the appropriate error
            if self.name.lower() in required_sections[self.version_num]:
                self.parse_errors.append(
                    RequiredSectionParseError(
                        "Couln't parse, raw section data is only one line or "
                        "less."
                    )
                )
                self.parse_tbs.append(traceback.format_exc())
            else:
                self.parse_errors.append(
                    SectionParseError(
                        "Couln't parse, raw section data is only one line or "
                        "less."
                    )
                )
                self.parse_tbs.append(traceback.format_exc())

    def validate(self):
        """
        Validates parsed sections from a LAS file depending on the section
        name, type and the file's version.

        This function checks the validity of parsed sections from a LAS
        file based on their name and type. It calls specific validation
        functions (validate_version, validate_well) for version and well
        headers and checks for read_errors attribute in data section.

        Parameters:
        ----------
        None

        Returns:
        -------
        None
        """

        if self.name == 'version' and self.type == 'header':
            errors = validate_version(self.parsed_section, self.version_num)
            if errors != []:
                for error in errors:
                    self.validate_errors.append(error)
        elif self.name == 'well' and self.type == 'header':
            errors = validate_well(self.parsed_section, self.version_num)
            if errors != []:
                for error in errors:
                    self.validate_errors.append(error)
        elif self.name == 'curves' and self.type == 'header':
            errors = validate_curves(self.parsed_section, self.version_num)
            if errors != []:
                for error in errors:
                    self.validate_errors.append(error)
        elif self.name == 'data' and self.type == 'data':
            if hasattr(self.parsed_section, 'read_errors'):
                for error in self.parsed_section.read_errors:
                    self.validate_errors.append(error)
        elif '_definition' in self.name and self.type == 'header':
            errors = validate_curves(self.parsed_section, self.version_num)
            if errors != []:
                for error in errors:
                    self.validate_errors.append(error)
        elif '_data' in self.name and self.type == 'data':
            if hasattr(self.parsed_section, 'read_errors'):
                for error in self.parsed_section.read_errors:
                    self.validate_errors.append(error)

    def __repr__(self):
        return (
            f"<LASSection(name={self.name!r}, type={self.type!r}, "
            f"version_num={self.version_num!r})>"
        )

    def __str__(self):
        s = (
            "LASSection\n"
            f"    Name: {self.name}\n"
            f"    Type: {self.type}\n"
            f"    Version: {self.version_num}\n"
            f"    Delimiter: {self.delimiter}\n"
        )
        if not hasattr(self, 'parse_error'):
            s += "    Parsed: True\n"
        else:
            s += (
                "    Parsed: False\n"
                "    Errors:\n"
                f"        Parsing Error: {self.parse_errors}\n"
                f"        Traceback:\n{self.parse_tbs}\n"
            )
        if not hasattr(self, 'validate_errors'):
            s += f"    Validated: {self.validated}\n"
        else:
            s += (
                f"    Validated: {self.validated}\n"
                "    Errors:\n"
                f"        Validation Error: {self.validate_errors}\n"
                f"        Traceback:\n{self.validate_tb}\n"
            )
        if hasattr(self, 'df'):
            s += f"    Rows: {len(self.df)}\n"

        return s

    def add_validate_errors(self, error, tb=None):
        if not hasattr(self, 'validate_errors'):
            self.validate_errors = []
        if not hasattr(self, 'validate_tb'):
            self.validate_tb = []
        self.validate_errors.append(error)
        if tb is not None:
            self.validate_tb.append(tb)


# Check for definition/curve and data column congruency
def check_definitions_and_format_data(def_section, data_section):
    """
    Checks the congruency between the definitions (header) and data
    sections of a LAS file and formats the data section accordingly.

    Parameters:
    -----------
    def_section : LASSection
        An instance of LASSection representing the definition section
        (usually the header) of a LAS file.

    data_section : LASSection
        An instance of LASSection representing the data section of a
        LAS file.

    Returns:
    --------
    None

    Note:
    ----
    If the number of rows in the definition section matches the number
    of columns in the data section, it renames the columns of the data
    section with the column names of the definition section. This
    operation modifies the data_section in-place.
    """
    def_rows = def_section.df.shape[0]
    # print(def_rows)
    data_cols = data_section.df.shape[1]
    # print(data_cols)
    if hasattr(def_section, 'df') and hasattr(data_section, 'df'):
        def_rows = def_section.df.shape[0]
        # print(def_rows)
        data_cols = data_section.df.shape[1]
        # print(data_cols)
        if def_rows == data_cols:
            data_section.df.rename(
                columns=dict(zip(
                        data_section.df.columns,
                        def_section.df.columns
                )),
                inplace=True
            )


class LASFile():
    """
    Class representing a Log ASCII Standard (LAS) file.

    This class provides methods for reading an LAS file, parsing its
    sections, validating those sections, and handling errors that occur
    during these processes.

    Attributes:
    ----------
    file_path : str
        The path of the LAS file.

    always_try_split : bool
        Whether or not to try to split the sections of the LAS file even
        when errors occur.

    sections : list of LASSection
        The sections of the LAS file, parsed and validated.

    version : LASSection
        The version section of the LAS file.

    version_num : str
        The version number of the LAS file.

    wrap : str
        The wrap of the LAS file.

    delimiter : str
        The delimiter used in the LAS file.

    read_error : str
        Error encountered during reading the file, if any.

    read_tb : str
        Traceback information for the read error, if any.

    open_error : str
        Error encountered during opening the file, if any.

    open_tb : str
        Traceback information for the open error, if any.

    version_error : str
        Error encountered during extraction of the version section,
        if any.

    version_tb : str
        Traceback information for the version error, if any.

    split_error : str
        Error encountered during splitting of the sections, if any.

    parse_error : dict
        Errors encountered during parsing of the sections, if any.

    validate_errors : dict
        Errors encountered during validation of the sections, if any.

    errors : dict
        All errors encountered during the processing of the LAS file.

    Methods:
    -------
    __init__(self, file_path=None, always_try_split=False)
        Initializes the LASFile object, reads the file, parses and validates
        its sections.

    read_file(self, file_path)
        Attempts to read the LAS file and handle any errors that occur during
        this process.

    get_version(self, data)
        Attempts to extract the version, wrap, and delimiter from the LAS file
        data.

    get_sections(self, data)
        Attempts to split the LAS file data into sections.

    parse_and_validate_sections(self, sections_dict)
        Attempts to parse and validate the sections of the LAS file.

    ensure_curve_and_data_congruency(self)
        Checks if the definition/curve and data columns of the LAS file are
        congruent.

    get_api(self)
        Attempts to extract the API number from the LAS file.

    __str__(self)
        Returns a user-friendly string representation of the LASFile object,
        including any errors that occurred.
    """
    def __init__(self, file_path=None):
        if file_path is not None:
            self.file_path = file_path
            self.sections = []
            self.curve_names = None
            # Try to open and read the file into a string
            data = self.read_file(self.file_path)

            if data is not None:
                # If it read correctly try to extract the version,
                # wrap, and delimiter and append it to the sections
                self.get_version(data)

                sections_dict = self.get_sections(data)

                self.parse_and_validate_sections(sections_dict)

            self.set_error_attributes()

    def read_file(self, file_path):
        # Try to open the file
        try:
            with open(self.file_path, 'r') as f:
                try:
                    data = f.read()
                    return data
                except Exception as e:
                    self.read_error = LASFileCriticalError(
                        f"Couldn't read file: {str(e)}"
                    )
                    tb = traceback.format_exc()
                    self.read_tb = f"Couldn't read file: {tb}"
                    return
        except FileNotFoundError:
            self.open_error = LASFileOpenError(
                f"File not found: {self.file_path}"
            )
            tb = traceback.format_exc()
            self.open_tb = f"File not found: {tb}"
            return
        except Exception as e:
            self.open_error = LASFileOpenError(
                f"Error opening file: {str(e)}"
            )
            tb = traceback.format_exc()
            self.open_tb = f"Error opening file: {tb}"
            return

    def get_version(self, data):
        # If it read correctly try to extract the version,
        # wrap, and delimiter and append it to the sections
        try:
            self.version = get_version_section(data)
            self.version_num = self.version.version_num
            self.wrap = self.version.wrap
            self.delimiter = self.version.delimiter
            self.sections.append(self.version)
            return
        # If a version couldn't be extracted, set the version error and
        # traceback, and return
        except Exception as e:
            self.version_error = LASFileCriticalError(
                f"Couldn't get version: {str(e)}"
            )
            tb = traceback.format_exc()
            self.version_tb = LASFileCriticalError(
                f"Couldn't get version: {tb}"
            )
            self.version_num = None
            return

    def get_sections(self, data):
        # Try to split the file into sections
        # Test if a version number was extracted and associated with
        # the LASFile object
        if hasattr(self, 'version_num'):
            # If a version number is associated with the LASFile object,
            # test if it is not None or an empty string
            if (
                getattr(self, 'version_num') is not None and
                getattr(self, 'version_num') != ''
            ):
                # If the version number is not None or an empty string,
                # try to split the file into sections
                try:
                    s = split_sections(data, self.version_num)
                except Exception as e:
                    # If an error occurs during the splitting of the
                    # sections, set the split error and traceback, and
                    # return
                    self.split_error = LASFileSplitError(
                        f"Couldn't split into sections: {str(e)}"
                    )
                    self.split_tb = traceback.format_exc()
                    return
                # If the sections were split correctly, check that the
                # minimum required sections are present and not empty
                # then return the sections dictionary
                if (
                    'version' in s.keys() and
                    'well' in s.keys() and
                    'curves' in s.keys() and
                    'data' in s.keys()
                ):
                    if (
                        s['version'] != '' and s['version'] is not None and
                        s['well'] != '' and s['well'] is not None and
                        s['curves'] != '' and s['curves'] is not None and
                        s['data'] != '' and s['data'] is not None
                    ):
                        return s
                    else:
                        self.split_error = LASFileSplitError(
                            "Couldn't split into minimum required "
                            "sections: Version, Well, Curves or Data "
                            "sections is empty."
                        )
                        return
                else:
                    self.split_error = LASFileSplitError(
                        "Couldn't split into minimum required "
                        "sections: Version, Well, Curves or Data "
                        "sections is missing."
                    )
                    return
        else:
            self.split_error = LASFileSplitError(
                "Couldn't split into minimum required "
                "sections: Version Number is missing."
            )
            return

    def parse_and_validate_sections(self, sections_dict):
        """
        Parse and validate the sections of a LAS file.

        Parameters
        ----------
        sections_dict : dict
            A dictionary containing the sections of the LAS file.

        Returns
        -------
        None
        """
        # If it split correctly and therefore the sections_dict is not
        # None, then generate las section items from the strings
        if sections_dict is not None:
            # Get the section name and raw text data from sections_dict
            for name, raw_data in sections_dict.items():
                # Skip the version section, it should already be parsed
                if name == 'version':
                    continue
                # Get the section type
                if (
                    name in header_section_names or
                    '_parameters' in name or
                    '_definition' in name
                ):
                    section_type = 'header'
                elif (
                    name in data_section_names or
                    '_data' in name
                ):
                    section_type = 'data'
                else:
                    section_type = ''
                # check if the sections attribute exists
                if hasattr(self, 'sections'):
                    # If it exists, check that the curves section is in
                    # the list of sections stored in the sections attribute
                    for section in self.sections:
                        if section.name == 'curves':
                            curves_section = section
                            if error_check(curves_section):
                                if hasattr(curves_section, 'df'):
                                    self.curve_names = (
                                        curves_section.
                                        df['mnemonic'].tolist()
                                    )
                # Try to create the section
                try:
                    section = LASSection(
                        name,
                        raw_data,
                        section_type,
                        self.version_num,
                        self.wrap,
                        delimiter=self.delimiter,
                        curve_names=self.curve_names
                    )
                    self.sections.append(section)
                except Exception as e:
                    # Try to create the section anyway without parsing
                    # or validating
                    try:
                        section = LASSection(
                            name,
                            raw_data,
                            section_type,
                            self.version_num,
                            self.wrap,
                            delimiter=self.delimiter,
                            parse_on_init=False,
                            validate_on_init=False,
                            curve_names=self.curve_names
                        )
                        section.parse_errors = [e]
                        self.sections.append(section)
                    except Exception as e:
                        if hasattr(self, 'parse_errors'):
                            self.parse_errors[name] = e
                        else:
                            self.parse_errors = {name: e}

            # Set LASFile section attributes
            for section in self.sections:
                # Set each section as an attribute of the LASFile
                setattr(self, section.name, section)

        # Run the function to ensure the curve and data sections
        # have the same number of curve mnemonics/data columns
        self.ensure_curve_and_data_congruency()

        # If the version is 3.0, add mnemonics to data sections
        if self.version_num == '3.0':
            self.add_mnemonics_to_data_sections()

        # Aggregate parse_errors and validate_errors from all sections
        self.aggregate_section_errors()
        return

    def ensure_curve_and_data_congruency(self):
        """Check for definition/curve and data column congruency"""
        # Check that the curves and data sections exist
        if hasattr(self, "curves") and hasattr(self, "data"):
            # Check that the curves and data sections were parsed
            # correctly into dataframes
            if (
                hasattr(getattr(self, "curves"), 'df') and
                hasattr(getattr(self, "data"), 'df')
            ):
                # Get the number of rows/curve definitions and the
                # number of columns/data points
                def_rows = getattr(self, "curves").df.shape[0]
                curves_df = getattr(self, "curves").df
                data_cols = getattr(self, "data").df.shape[1]
                # If the number of rows/curve definitions in the
                # definition section matches the number of columns in
                # the data section, rename the columns of the data
                # section to the curve mnemonics
                if def_rows == data_cols:
                    # Check if there are repeated curve mnemonics
                    if (
                        not len(
                            curves_df.mnemonic.unique()
                        ) == len(
                            curves_df.mnemonic
                        )
                    ):
                        # Get a list of the mnemonics that are repeated
                        repeated_mnemonics = curves_df.mnemonic[
                            curves_df.mnemonic.duplicated(keep=False)
                        ].unique()
                        # For each unique repeated mnemonic make a df of
                        # the repeated mnemonics
                        for mnemonic in repeated_mnemonics:
                            # Get a df of the repeated mnemonics
                            repeated_mnemonics_df = (
                                curves_df.loc[
                                    curves_df['mnemonic'] == mnemonic
                                ]
                            )
                            # Reset the index of the repeated mnemonics df
                            repeated_mnemonics_df.reset_index(inplace=True)
                            # for each row in the repeated mnemonics df,
                            # append an underscore and the index digit
                            # to the end of the mnemonic, except the
                            # first instance of the repeated mnemonic
                            # format: {old_index: [new_mnemonic]}
                            new_repeated_mnemonics = {}
                            for index, row in repeated_mnemonics_df.iterrows():
                                if index == 0:
                                    new_repeated_mnemonics[row['index']] = (
                                        row['mnemonic']
                                    )
                                else:
                                    new_repeated_mnemonics[row['index']] = (
                                        f"{row['mnemonic']}_{index}"
                                    )

                            # Create a copy of the curves df
                            new_curves_df = curves_df
                            # replace the old mnemonics with the new
                            # ones by index
                            for index, new_mnemonic in \
                                    new_repeated_mnemonics.items():
                                new_curves_df.loc[index, 'mnemonic'] = (
                                    new_mnemonic
                                )
                            # Replace the curves df with the new one
                            setattr(
                                getattr(self, 'curves'),
                                'df',
                                new_curves_df)
                    # Rename the columns of the data section
                    getattr(self, "data").df.rename(
                        columns=dict(zip(
                            getattr(self, "data").df.columns,
                            getattr(self, "curves").df.mnemonic.values
                        )),
                        inplace=True
                    )
                # If the number of rows/curve definitions in the
                # definition section does not match the number of
                # columns in the data section, set a validation error
                else:
                    if not hasattr(getattr(self, 'curves'), 'validate_errors'):
                        setattr(getattr(self, 'curves'), 'validate_errors', [])
                    if not hasattr(getattr(self, 'data'), 'validate_errors'):
                        setattr(getattr(self, 'data'), 'validate_errors', [])
                    getattr(self, 'curves').validate_errors.append(
                        LASFileCriticalError(
                            "Curves and data sections are not "
                            "congruent."
                        )
                    )
                    getattr(self, 'data').validate_errors.append(
                        LASFileCriticalError(
                            "Curves and data sections are not "
                            "congruent."
                        )
                    )

    def add_mnemonics_to_data_sections(self):
        """
        Adds mnemonics from associated definition sections to data sections
        for LAS version 3.0 files.
        """
        # Loop through sections
        for section in self.sections:
            # If the section is a data section
            if section.type == 'data':
                # and has an association
                if hasattr(section, 'association'):
                    # and association is not None
                    if section.association is not None:
                        # Map 'curve' to 'curves' if needed
                        association = section.association
                        if association == 'curve':
                            association = 'curves'

                        try:
                            # Get the mnemonics from the associated definition
                            # section
                            mnemonics = getattr(
                                getattr(self, association),
                                'df'
                            )['mnemonic'].tolist()

                            # Check if number of mnemonics matches columns
                            if len(mnemonics) != len(section.df.columns):
                                data_cols = len(section.df.columns)
                                mnem_cols = len(mnemonics)
                                error_msg = (
                                    f"Columns/mnemonics mismatch: Data has "
                                    f"{data_cols} columns, but {association} "
                                    f"section has {mnem_cols} mnemonics"
                                )
                                # Create or append to section validation errors
                                if not hasattr(section, 'validate_errors'):
                                    section.validate_errors = []
                                # Add as a critical error since data integrity
                                # can't be guaranteed
                                section.validate_errors.append(
                                    LASFileCriticalError(error_msg)
                                )
                                # Don't try to fix or apply column names
                            else:
                                # Add the mnemonics to the data section only if
                                # they match
                                section.df.columns = mnemonics

                        except Exception as e:
                            # Catch any other errors that might occur
                            if not hasattr(section, 'validate_errors'):
                                section.validate_errors = []
                            section.validate_errors.append(
                                LASFileCriticalError(
                                    f"Error adding mnemonics to data section: "
                                    f"{str(e)}"
                                )
                            )
                            # Don't try to assign generic column names

    def aggregate_section_errors(self):
        if hasattr(self, 'sections'):
            for section in self.sections:
                # Aggregate parse_errors
                # If the current section has a parse error
                if hasattr(section, 'parse_errors'):
                    # Instantiate the parse_errors attribute if it
                    # doesn't exist
                    if not hasattr(self, 'parse_errors'):
                        setattr(self, 'parse_errors', {})
                    # Add the current parse error for the current
                    # section to the LASFile parse_errors dictionary
                    self.parse_errors[section.name] = (
                            section.parse_errors
                        )

                # Aggregate validate_errors
                # If the current section has validate errors
                if hasattr(section, 'validate_errors'):
                    # Instantiate the validate_errors attribute if it
                    # doesn't exist
                    if not hasattr(self, 'validate_errors'):
                        setattr(self, 'validate_errors', {})
                    # Add the current validate error for the current
                    # section to the LASFile validate_errors dictionary
                    getattr(self, 'validate_errors')[section.name] = (
                            section.validate_errors
                        )

    def set_error_attributes(self):
        # Aggregate all the errors into one dictionary and store it
        # in lasfile.errors
        self.errors = {}
        if hasattr(self, 'open_error'):
            self.errors['open_error'] = getattr(self, 'open_error')
        if hasattr(self, 'read_error'):
            self.errors['read_error'] = getattr(self, 'read_error')
        if hasattr(self, 'split_error'):
            self.errors['split_error'] = getattr(self, 'split_error')
        if hasattr(self, 'version_error'):
            self.errors['version_error'] = getattr(self, 'version_error')
        if hasattr(self, 'parse_errors'):
            self.errors['parse_errors'] = getattr(self, 'parse_errors')
        if hasattr(self, 'validate_errors'):
            self.errors['validate_errors'] = getattr(self, 'validate_errors')

    def __str__(self):
        s = f"LASFile: {self.file_path}\n"
        if hasattr(self, 'open_error'):
            s += f"Open Error: {self.open_error}\n"
        if hasattr(self, 'read_error'):
            s += f"Reading Error: {self.read_error}\n"
        if hasattr(self, 'version_error'):
            s += f"Version Extraction Error: {self.version_error}\n"
        if hasattr(self, 'split_error'):
            s += f"Section Splitting Error: {self.split_error}\n"
        if hasattr(self, 'parse_errors'):
            s += f"Parsing Error: {self.parse_errors}\n"
        if hasattr(self, 'validate_errors'):
            s += f"Validation Error: {getattr(self, 'validate_errors')}"
        if hasattr(self, 'sections'):
            for section in self.sections:
                s += str(f"  {section}")
        return s

    def get_api(self):
        # If the las file has a well section, try to get the api from it
        if hasattr(self, 'well'):
            try:
                # Check if 'UWI', 'uwi', 'API', or 'api' is present in the
                # 'mnemonic' column
                mask = getattr(self, "well").df['mnemonic'].str.lower().isin(
                    ['uwi', 'api']
                )
                # Filter the DataFrame using the mask
                filtered_df = getattr(self, "well").df[mask]
                # Get the corresponding values for the matched mnemonics
                matched_values = filtered_df['value'].tolist()

                # Attempt to load all matched values into an APINumber
                # objects
                valid_values = []
                for value in matched_values:
                    try:
                        APINumber(value)
                        valid_values.append(value)
                    except Exception:
                        pass

                # Remove Null values from valid_values
                valid_values = [x for x in valid_values if x is not None]

                # If there are matched values, check if they have the same
                # first 10 characters
                if len(valid_values) > 0:
                    if all(
                        APINumber(x).unformatted_10_digit == APINumber(
                            valid_values[0]
                            ).unformatted_10_digit
                        for x in valid_values
                    ):
                        return APINumber(valid_values[0])
                    else:
                        # If they don't have the same first 10 characters,
                        # return the longest valid_value as an APINumber
                        return APINumber(max(valid_values, key=len))
                else:
                    return None
            except Exception as e:
                raise e
        else:
            return None


def read(fp):
    """
    Read a LAS file and return a LASFile object

    Parameters
    ----------
    las_file : str
        Path to the LAS file

    Returns
    -------
    LASFile object
    """
    return LASFile(file_path=fp)


def arrange(section, min_spaces=2, header=True):
    """
    Arrange the lines in a section to be neatly formatted, based on the
    section type ('header' or 'data').

    Parameters
    ----------
    section (LASSection):
        The LASSection object with attributes `type` and `df`
        (DataFrame).
    min_spaces (int):
        Minimum spaces between columns. Default is 2. Minimum value
        allowed is 1.
    header (bool):
        Whether to include the header and separator lines. Default is
        True.

    Returns
    -------
    str:
        Formatted text block ready for writing to a file.
    """
    # Ensure min_spaces is at least 1
    min_spaces = max(min_spaces, 1)

    # Extract the DataFrame from the section
    df = section.df

    # Handle formatting based on section type
    if section.type == 'header':
        # Calculate the maximum widths for each column, considering
        # minimum header lengths
        # Minimum width to fit "MNEM"
        max_mnem_width = max(df['mnemonic'].str.len().max(), 4)
        # Minimum width to fit "UNIT"
        max_unit_width = max(df['units'].fillna('').str.len().max(), 4)
        # Minimum width to fit "VALUE"
        max_value_width = max(
            df['value'].apply(lambda x: len(str(x))).max(), 5
        )
        # Minimum width to fit "DESCRIPTION"
        max_desc_width = max(
            df['description'].fillna('').str.len().max(), 11
        )

        # Define header and separator lines based on calculated widths
        header_line = (
            f"# MNEM.UNIT{' ' * (max_unit_width - 4 + min_spaces)} "
            f"VALUE{' ' * (max_value_width - 5 + min_spaces)}"
            f"DESCRIPTION"
        )
        separator = (
            f"# {'-' * max_mnem_width}.{'-' * max_unit_width}"
            f"{' ' * min_spaces} {'-' * max_value_width}"
            f"{' ' * min_spaces}{'-' * max_desc_width}"
        )

        # Initialize formatted lines list
        formatted_lines = []

        # Add headers if header is True
        if header:
            formatted_lines.extend([header_line, separator])

        # Format each line according to the calculated widths
        for _, row in df.iterrows():
            mnemonic = row['mnemonic']
            unit = row['units'] if row['units'] else ''
            value = str(row['value']) if row['value'] else ''
            description = row['description'] if row['description'] else ''

            line = (
                f"  {mnemonic:<{max_mnem_width}}"
                f".{unit:<{max_unit_width}}{' ' * min_spaces} "
                f"{value:<{max_value_width}}{' ' * min_spaces}"
                f": {description}"
            )
            formatted_lines.append(line)

    elif section.type == 'data':
        # Calculate maximum widths dynamically based on the DataFrame's
        # headers
        max_col_widths = {
            col: max(df[col].astype(str).apply(len).max(), len(col))
            for col in df.columns
        }

        # Define dynamic header and separator lines based on DataFrame
        # headers
        header_line = "# " + "  ".join(
            [f"{col:<{max_col_widths[col]}}" for col in df.columns]
        )
        separator = "# " + "  ".join(
            ['-' * max_col_widths[col] for col in df.columns]
        )

        # Initialize formatted lines list
        formatted_lines = []

        # Add headers if header is True
        if header:
            formatted_lines.extend([header_line, separator])

        # Format each row of data based on column widths
        for _, row in df.iterrows():
            line = "  " + "  ".join(
                [
                    f"{str(row[col]):<{max_col_widths[col]}}"
                    for col in df.columns
                ]
            )
            formatted_lines.append(line)

    # Combine all lines into a single text block and add a newline at the end
    return "\n".join(formatted_lines) + "\n"


def write(lf, overwrite=False, file_path=None, version=None):
    """
    Write a LASFile object to a LAS file

    Parameters
    ----------
    lf : LASFile object
        The LASFile object to write to a file

    overwrite : bool
        If True, overwrite the file if it already exists

    file_path : str
        The path to write the file to, if this is None, it will write to
        the path stored in the LASFile object, if overwrite is True, it
        will overwrite the file at the path stored in the LASFile object
    """
    # Check if the file_path is None, if it is, set it to the file_path
    # attribute of the LASFile object
    if file_path is None:
        file_path = lf.file_path
    if file_path is None:
        raise ValueError("file_path is None")
    # Check if the file already exists and if overwrite is False, raise
    # a FileExistsError
    if os.path.exists(file_path) and not overwrite:
        raise FileExistsError(
            "The file already exists, set overwrite to True to overwrite"
        )
    # Check if a version is provided, if it is not, use the version
    # attribute of the LASFile object
    if version is None:
        version = lf.version.version_num
    # Check for errors in the LASFile object then write the file if
    # there are no errors
    if error_check(lf):
        with open(file_path, 'w') as f:
            # Write the version section
            f.write("~Version\n")
            f.write(
                f"  VERS.    {version}:    "
                f"CWLS LOG ASCII STANDARD -VERSION {version}\n"
            )
            f.write(
                "  WRAP.     NO:    ONE LINE PER DEPTH STEP\n"
            )
            if version == '3.0':
                f.write(
                    "  DLM . COMMA : DELIMITING CHARACTER (SPACE TAB OR COMMA)"
                )
            # Write the well section
            f.write("~Well\n")
            # Arrange the well section
            well_section = arrange(lf.well)
            f.write(well_section)

            # Write the curves section
            f.write("~Curves\n")
            # Arrange the curves section
            curves_section = arrange(lf.curves)
            f.write(curves_section)

            # If the lasfile has a parameters section, write it
            if hasattr(lf, 'parameters'):
                # Write the parameters section
                f.write("~Parameters\n")
                # Arrange the parameters section
                parameters_section = arrange(lf.parameters)
                f.write(parameters_section)

            # If the lasfile has an other section, write it
            if hasattr(lf, 'other'):
                # Write the other section
                f.write(lf.other.raw_data)
                f.write("\n")

            # Write any other sections if it is version 3.0
            if version == '3.0':
                for section in lf.sections:
                    if section.name not in [
                        'version',
                        'well',
                        'parameters',
                        'other',
                        'curves',
                        'data'
                    ]:
                        f.write(f"~{section.name}\n")
                        section_section = arrange(section)
                        f.write(section_section)

            # Write the data section
            f.write("~ASCII Data\n")
            # Arrange the data section
            data_section = arrange(lf.data)
            f.write(data_section)


def api_from_las(input):
    # If the input is a string, assume it's a file path and try to read
    # it into a LASFile object
    if isinstance(input, str):
        try:
            las = read(input)
        except Exception as e:
            raise e
    elif isinstance(input, LASFile):
        las = input
    else:
        las = LASFile()
    # If the las has a well section, try to get the api from it
    if hasattr(las, 'well'):
        try:
            # Check if 'UWI', 'uwi', 'API', or 'api' is present in the
            # 'mnemonic' column
            mask = getattr(las, "well").df['mnemonic'].str.lower().isin(
                ['uwi', 'api']
            )
            # Filter the DataFrame using the mask
            filtered_df = getattr(las, "well").df[mask]
            # Get the corresponding values for the matched mnemonics
            matched_values = filtered_df['value'].tolist()

            # Attempt to load all matched values into an APINumber
            # objects
            valid_values = []
            for value in matched_values:
                try:
                    APINumber(value)
                    valid_values.append(value)
                except Exception:
                    pass

            # Remove Null values from valid_values
            valid_values = [x for x in valid_values if x is not None]

            # If there are matched values, check if they have the same
            # first 10 characters
            if len(valid_values) > 0:
                if all(
                    APINumber(x).unformatted_10_digit == APINumber(
                        valid_values[0]
                        ).unformatted_10_digit
                    for x in valid_values
                ):
                    return APINumber(valid_values[0])
                else:
                    # If they don't have the same first 10 characters,
                    # return the longest valid_value as an APINumber
                    return APINumber(max(valid_values, key=len))
            else:
                return None
        except Exception as e:
            raise e
    else:
        return None


def error_check(las, critical_only=True):
    """
    Check if a LASFile object has any errors

    Parameters
    ----------
    las : LASFile object
        The LASFile object to check for errors
    critical_only : bool
        If True, only check for critical errors, otherwise check for all
        errors

    Returns
    -------
    bool
        True if no errors are found, False if any errors are found
    """
    # Check if an las object, either LASSection or LASFile, has any
    # errors and if critical_only is True, if any are of type
    # LASFileCriticalError return False when one is found,
    # otherwise return False when any error is found.
    if critical_only:
        # Check if the las object has an open_error
        if hasattr(las, 'open_error'):
            if isinstance(las.open_error, LASFileCriticalError):
                return False
        # Check if the las object has a read_error
        if hasattr(las, 'read_error'):
            if type(las.read_error) is dict:
                for sec_name, error in las.read_error.items():
                    if isinstance(error, LASFileCriticalError):
                        return False
            elif type(las.read_error) is Exception:
                if isinstance(las.read_error, LASFileCriticalError):
                    return False
        # Check if the las object has a version_error
        if hasattr(las, 'version_error'):
            if type(las.version_error) is dict:
                for sec_name, error in las.version_error.items():
                    if isinstance(error, LASFileCriticalError):
                        return False
            elif type(las.version_error) is Exception:
                if isinstance(las.version_error, LASFileCriticalError):
                    return False
        # Check if the las object has a split_error
        if hasattr(las, 'split_error'):
            if isinstance(las.split_error, LASFileCriticalError):
                return False
        # Check if the las object has a parse_errors
        if hasattr(las, 'parse_errors'):
            if type(las.parse_errors) is dict:
                for sec_name, error in las.parse_errors.items():
                    if isinstance(error, LASFileCriticalError):
                        return False
            elif type(las.parse_errors) is Exception:
                if isinstance(las.parse_errors, LASFileCriticalError):
                    return False
        # Check if the las object has a validate_errors
        if hasattr(las, 'validate_errors'):
            if type(las.validate_errors) is dict:
                for sec_name, error_list in las.validate_errors.items():
                    for error in error_list:
                        if isinstance(error, LASFileCriticalError):
                            return False
            if type(las.validate_errors) is list:
                for error in las.validate_errors:
                    if isinstance(error, LASFileCriticalError):
                        return False
            if type(las.validate_errors) is Exception:
                if isinstance(las.validate_errors, LASFileCriticalError):
                    return False
        # If no critical errors are found, return True
        return True
    else:
        # Check if the las object has any errors
        if hasattr(las, 'open_error'):
            return False
        if hasattr(las, 'read_error'):
            return False
        if hasattr(las, 'version_error'):
            return False
        if hasattr(las, 'split_error'):
            return False
        if hasattr(las, 'parse_errors'):
            return False
        if hasattr(las, 'validate_errors'):
            return False
        return True

# END EMBEDDED LASFILE PARSER



# ---------------------------------------------------------------------------
# IDTW application. The embedded LAS reader above is third-party MIT code.
# This section implements Fang et al. (2021), eq. 4, with explicit engineering
# choices: optional robust scaling, equal curve weights, a band around the line
# joining interval endpoints, step penalties, and masked-gap traversal.
# No well-top snapping is applied: arbitrary extrema are not geological picks.
# ---------------------------------------------------------------------------
import argparse
import csv
import math
import queue
import sys
import threading
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from dataclasses import asdict, replace
import copy
import hashlib
import colorsys
import itertools
import time
import zipfile

import numpy as np
from openpyxl import load_workbook


@dataclass
class Well:
    path: str
    name: str
    uwi: str
    depth: np.ndarray
    curves: dict
    units: dict
    notes: list = field(default_factory=list)


@dataclass
class Marker:
    name: str
    well: str
    uwi: str
    md: float
    x: float | None = None
    y: float | None = None
    z: float | None = None


@dataclass
class Params:
    step: float = 0.5
    band: float = 50.0
    max_points: int = 2000
    gap: float = 3.0
    normalize: str = 'robust'
    k: int = 0
    penalty: float = 0.05
    ref_start: float | None = None
    ref_end: float | None = None
    target_start: float | None = None
    target_end: float | None = None


@dataclass
class Result:
    ref: Well
    target: Well
    curves: list
    params: Params
    zr: np.ndarray
    zt: np.ndarray
    xr: np.ndarray
    xt: np.ndarray
    path: np.ndarray
    mapped: np.ndarray
    similarity: np.ndarray
    rows: list
    stats: dict
    notes: list
    reference_markers: list = field(default_factory=list)
    control_markers: list = field(default_factory=list)
    provenance: dict = field(default_factory=dict)


class Cancelled(Exception):
    pass


class _EncodedLASFile(LASFile):
    def read_file(self, file_path):
        raw = Path(file_path).read_bytes()
        if len(raw) > 100 * 1024 * 1024:
            raise ValueError('LAS больше 100 МБ. Выберите файл меньшего размера.')
        if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
            content, self.source_encoding = raw.decode('utf-16'), 'utf-16'
        else:
            try:
                content, self.source_encoding = raw.decode('utf-8-sig'), 'utf-8'
            except UnicodeDecodeError:
                content, self.source_encoding = raw.decode('cp1251'), 'cp1251'
        # Upstream expects uppercase section initials and does not consistently
        # skip full-line comments in the numeric section.
        lines = []
        for line in content.splitlines():
            stripped = line.lstrip()
            if stripped.startswith('#'):
                continue
            if stripped.startswith('~') and len(stripped) > 1:
                line = '~' + stripped[1].upper() + stripped[2:]
            lines.append(line)
        return '\n'.join(lines) + '\n'


def _flatten_errors(value):
    if isinstance(value, dict):
        for item in value.values():
            yield from _flatten_errors(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _flatten_errors(item)
    elif value:
        yield value


def load_las(path):
    """Use the embedded bzlmnop LASFile reader, then validate numerical input."""
    path = Path(path).resolve()
    if path.stat().st_size > 100 * 1024 * 1024:
        raise ValueError('LAS больше 100 МБ. Выберите файл меньшего размера.')
    las = _EncodedLASFile(str(path))
    errors = []
    for attr in ('open_error', 'read_error', 'version_error', 'split_error',
                 'parse_errors', 'validate_errors'):
        errors.extend(_flatten_errors(getattr(las, attr, None)))
    critical = [e for e in errors if isinstance(e, LASFileCriticalError)]
    if critical:
        raise ValueError(f'{path.name}: ' + '; '.join(str(e) for e in critical[:4]))
    if getattr(las, 'version_num', None) not in ('1.2', '2.0', '3.0'):
        raise ValueError(f'{path.name}: не поддерживается версия LAS.')
    for section in ('data', 'curves', 'well'):
        obj = getattr(las, section, None)
        if obj is None or not isinstance(getattr(obj, 'df', None), pd.DataFrame):
            raise ValueError(f'{path.name}: не удалось прочитать секцию {section}.')
        if section == 'data' and (getattr(obj, 'read_errors', None) or getattr(obj, 'parse_errors', None)):
            raise ValueError(f'{path.name}: ошибки чтения числовой секции LAS.')
    data, metadata = las.data.df, las.curves.df
    if data.shape[0] < 8 or data.shape[1] < 2 or data.shape[1] != len(metadata):
        raise ValueError(f'{path.name}: нужны минимум 8 строк и глубина + одна кривая; '
                         'число кривых должно совпадать с ~Curve.')
    names = [_identifier(v).upper() for v in metadata['mnemonic']]
    if len(set(names)) != len(names) or any(not n for n in names):
        raise ValueError(f'{path.name}: пустые или повторяющиеся мнемоники.')
    try:
        values = data.apply(pd.to_numeric, errors='raise').to_numpy(dtype=float, copy=True)
    except (ValueError, TypeError) as error:
        raise ValueError(f'{path.name}: нечисловые значения в ~ASCII: {error}') from error
    wellmeta = {_identifier(row['mnemonic']).upper(): row['value'] for _, row in las.well.df.iterrows()}
    null = _number(wellmeta.get('NULL', -999.25), f'{path.name}: NULL')
    values[values == null] = np.nan
    values[~np.isfinite(values)] = np.nan
    depth_index = next((i for i, name in enumerate(names) if name in ('DEPT', 'DEPTH', 'MD')), None)
    if depth_index is None:
        raise ValueError(f'{path.name}: нужна кривая глубины DEPT, DEPTH или MD.')
    unit_values = [_identifier(u) for u in metadata['units']]
    depth_unit = unit_values[depth_index].strip().casefold()
    if depth_unit in ('m', 'metre', 'meter', 'meters', 'metres', 'м'):
        factor = 1.
    elif depth_unit in ('ft', 'f', 'feet', 'foot'):
        factor = .3048
    else:
        raise ValueError(f'{path.name}: единица глубины {depth_unit!r} не распознана; '
                         'укажите M или FT в заголовке LAS.')
    depth = values[:, depth_index] * factor
    if not np.all(np.isfinite(depth)):
        raise ValueError(f'{path.name}: в столбце глубины есть NULL/пропуски.')
    order = np.argsort(depth, kind='stable')
    depth, values = depth[order], values[order]
    if np.any(np.diff(depth) <= 0):
        raise ValueError(f'{path.name}: повторяющиеся MD. Объедините или удалите дубликаты до загрузки.')
    notes = [f'{path.name}: кодировка {las.source_encoding}; LAS {las.version_num}.']
    if not np.array_equal(order, np.arange(len(order))):
        notes.append(f'{path.name}: отсчёты отсортированы по MD.')
    if factor != 1:
        notes.append(f'{path.name}: глубина переведена из футов в метры; markers.xlsx ожидается в метрах.')
    for error in errors[:6]:
        notes.append(f'{path.name}: замечание LAS-парсера: {error}')
    curves = {name: values[:, i].copy() for i, name in enumerate(names) if i != depth_index}
    units = {name: unit_values[i] for i, name in enumerate(names) if i != depth_index}
    return Well(str(path), _identifier(wellmeta.get('WELL')) or path.stem,
                _identifier(wellmeta.get('UWI')) or _identifier(wellmeta.get('API')),
                depth, curves, units, notes)


def _identifier(value):
    if value is None:
        return ''
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(float(value)):
            return ''
        if float(value).is_integer():
            return str(int(value))
    return str(value).strip()


def _key(value):
    return _identifier(value).casefold()


def _number(value, context, optional=False):
    if value is None or str(value).strip() == '':
        if optional:
            return None
        raise ValueError(f'{context}: пустое числовое значение.')
    try:
        result = float(str(value).strip().replace(',', '.'))
    except (ValueError, TypeError):
        raise ValueError(f'{context}: ожидалось число, получено {value!r}.') from None
    if not np.isfinite(result):
        raise ValueError(f'{context}: число должно быть конечным.')
    return result


def load_markers(path):
    """Read the first sheet with the seven required headers in its first 20 rows.

    MD is measured depth in metres. X/Y/Z are preserved, not used as MD or TVD.
    Text UWI values retain leading zeroes. For numeric Excel cells only a simple
    zero-padding number format can recover displayed leading zeroes.
    """
    required = ('маркер', 'скважина', 'uwi', 'md', 'x', 'y', 'z')
    book = load_workbook(path, read_only=True, data_only=True)
    markers = []
    try:
        selected = None
        for sheet in book:
            for cells in sheet.iter_rows(min_row=1, max_row=20):
                headers = [_key(c.value) for c in cells]
                if all(h in headers for h in required):
                    if any(headers.count(h) != 1 for h in required):
                        raise ValueError('Повторяющиеся заголовки в markers.xlsx.')
                    selected = (sheet, cells[0].row, {h: headers.index(h) for h in required})
                    break
            if selected:
                break
        if selected is None:
            raise ValueError('В первых 20 строках листов не найдены столбцы: '
                             'Маркер, Скважина, UWI, MD, X, Y, Z.')
        sheet, header_row, columns = selected
        for cells in sheet.iter_rows(min_row=header_row + 1):
            values = [c.value for c in cells]
            if not any(v is not None and str(v).strip() for v in values):
                continue
            def get(name):
                return values[columns[name]] if columns[name] < len(values) else None
            name, well, uwi = (_identifier(get(h)) for h in required[:3])
            uwi_cell = cells[columns['uwi']]
            fmt = uwi_cell.number_format or ''
            if uwi.isdigit() and fmt and set(fmt) == {'0'}:
                uwi = uwi.zfill(len(fmt))
            if not name or not (well or uwi):
                raise ValueError(f'{sheet.title}, строка {cells[0].row}: '
                                 'нужны имя маркера и Скважина или UWI.')
            ctx = f'{sheet.title}, строка {cells[0].row}'
            nums = [_number(get(h), f'{ctx}, {h.upper()}', h != 'md')
                    for h in required[3:]]
            markers.append(Marker(name, well, uwi, *nums))
    finally:
        book.close()
    if not markers:
        raise ValueError('markers.xlsx не содержит маркеров.')
    seen, clean = {}, []
    for marker in markers:
        identity = ('uwi', _key(marker.uwi)) if marker.uwi else ('well', _key(marker.well))
        key = identity + (_key(marker.name),)
        if key in seen:
            if abs(seen[key].md - marker.md) > 1e-6:
                raise ValueError(f'Конфликт: маркер {marker.name!r} для '
                                 f'{marker.uwi or marker.well!r} имеет несколько MD.')
            continue
        seen[key] = marker
        clean.append(marker)
    return clean


def marker_groups(markers):
    buckets = {}
    for marker in markers:
        identity = ('UWI', _key(marker.uwi)) if marker.uwi else ('WELL', _key(marker.well))
        buckets.setdefault(identity, []).append(marker)
    result = {}
    for (kind, identity), values in buckets.items():
        first = values[0]
        label = f'{first.well or "Без имени"} | {kind}: {first.uwi or first.well}'
        result[label] = sorted(values, key=lambda m: m.md)
    return result


def match_group(well, groups):
    # UWI is authoritative; do not fall back to an unrelated UWI on name match.
    if well.uwi:
        candidates = [k for k, ms in groups.items()
                      if any(_key(m.uwi) == _key(well.uwi) for m in ms)]
        if len(candidates) == 1:
            return candidates[0]
    names = {_key(well.name), _key(Path(well.path).stem)} - {''}
    candidates = [k for k, ms in groups.items() if any(
        _key(m.well) in names and not (well.uwi and m.uwi and _key(well.uwi) != _key(m.uwi))
        for m in ms)]
    return candidates[0] if len(candidates) == 1 else ''


def _runs(mask):
    edges = np.diff(np.r_[False, mask, False].astype(np.int8))
    return zip(np.flatnonzero(edges == 1), np.flatnonzero(edges == -1))


def resample_masked(depth, values, grid, max_gap):
    """Interpolate only between adjacent finite source samples, never across NULL.

    Missing acquisition intervals wider than max_gap also stay missing. Exact
    valid observations remain valid even when adjacent intervals are too wide.
    """
    result = np.full(len(grid), np.nan)
    right = np.searchsorted(depth, grid, side='left')
    clipped = np.clip(right, 0, len(depth) - 1)
    exact = np.isclose(depth[clipped], grid, rtol=0, atol=1e-7)
    result[exact] = values[clipped[exact]]
    between = (~exact) & (right > 0) & (right < len(depth))
    pos = np.flatnonzero(between)
    hi, lo = right[pos], right[pos] - 1
    good = (np.isfinite(values[lo]) & np.isfinite(values[hi]) &
            (depth[hi] - depth[lo] <= max_gap + 1e-9))
    pos, lo, hi = pos[good], lo[good], hi[good]
    fraction = (grid[pos] - depth[lo]) / (depth[hi] - depth[lo])
    result[pos] = values[lo] + fraction * (values[hi] - values[lo])
    # Keep acquisition breaks even if a coarser grid jumps entirely over them.
    # Mask neighbouring grid samples conservatively, so Hilbert cannot bridge
    # a source NULL that happens to fall between two resampled grid points.
    large = np.flatnonzero(np.diff(depth) > max_gap + 1e-9)
    barriers = np.r_[depth[~np.isfinite(values)], (depth[large]+depth[large+1])/2]
    barriers = barriers[(barriers >= grid[0]) & (barriers <= grid[-1])]
    if len(barriers):
        locations = np.searchsorted(grid, barriers)
        result[np.clip(locations, 0, len(grid)-1)] = np.nan
        result[np.clip(locations-1, 0, len(grid)-1)] = np.nan
    return result


def analytic_signal(values):
    """FFT Hilbert analytic signal, equivalent to the standard frequency mask."""
    n = len(values)
    multipliers = np.zeros(n)
    if n:
        multipliers[0] = 1
        if n % 2 == 0:
            multipliers[n // 2] = 1
            multipliers[1:n // 2] = 2
        else:
            multipliers[1:(n + 1) // 2] = 2
    return np.fft.ifft(np.fft.fft(values) * multipliers)


def analytic_masked(values):
    """Transform continuous valid runs separately with reflected edge padding.

    Reflection reduces finite-record discontinuities; it is an implementation
    choice, not a claim that Hilbert edge artefacts disappear.
    """
    result = np.full(len(values), np.nan + 1j * np.nan, dtype=complex)
    for start, stop in _runs(np.isfinite(values)):
        if stop - start < 8:
            continue
        segment = values[start:stop]
        padding = min(128, len(segment) - 1)
        padded = np.pad(segment, padding, mode='reflect')
        result[start:stop] = analytic_signal(padded)[padding:-padding]
    return result


def normalize_curve(values, mode):
    result = values.copy()
    valid = np.isfinite(values)
    if valid.sum() < 8:
        raise ValueError('Менее 8 действительных отсчётов выбранной кривой.')
    q = values[valid]
    if np.ptp(q) <= 1e-10 * max(1., float(np.max(np.abs(q)))):
        raise ValueError('Выбранная кривая постоянна: сопоставление неоднозначно.')
    if mode == 'robust':
        center = float(np.median(q))
        scale = float(np.quantile(q, .75) - np.quantile(q, .25)) / 1.349
        if scale <= 1e-12:
            scale = float(np.std(q))
        result[valid] = (q - center) / scale
    return result


def _grid(well, start, end, step, limit):
    start = well.depth[0] if start is None else start
    end = well.depth[-1] if end is None else end
    if not (well.depth[0] <= start < end <= well.depth[-1]):
        raise ValueError(f'{well.name}: интервал должен лежать в '
                         f'[{well.depth[0]:.3f}; {well.depth[-1]:.3f}] м.')
    count = min(limit, int(math.ceil((end - start) / step)) + 1)
    if count < 8:
        raise ValueError('Интервал слишком короткий: нужно хотя бы 8 отсчётов.')
    return np.linspace(start, end, count)


def _cost_row(a, b, i, js, half_window):
    """Equal mean of available curve costs; ratio of window energy sums.

    A curve contributes only when its entire requested window is valid in both
    wells. No epsilon converts zero-energy data into a perfect match.
    """
    numerator = np.zeros((len(js), a.shape[1]))
    denominator = np.zeros_like(numerator)
    valid = np.ones_like(numerator, dtype=bool)
    for offset in range(-half_window, half_window + 1):
        ri, tj = i + offset, js + offset
        if ri < 0 or ri >= len(a):
            valid[:] = False
            continue
        inside = (tj >= 0) & (tj < len(b))
        av = a[ri]
        bv = b[np.clip(tj, 0, len(b) - 1)]
        good = np.isfinite(av)[None, :] & np.isfinite(bv) & inside[:, None]
        valid &= good
        numerator += np.where(good, np.abs(av - bv) ** 2, 0)
        denominator += np.where(good, 2 * (np.abs(av) ** 2 + np.abs(bv) ** 2), 0)
    valid &= denominator > np.finfo(float).tiny
    costs = np.divide(numerator, denominator, out=np.zeros_like(numerator), where=valid)
    count = valid.sum(axis=1)
    average = np.divide((costs * valid).sum(axis=1), count,
                        out=np.full(len(js), 1.25), where=count > 0)
    return np.clip(average, 0, 1.25), count > 0


def _informative(values, index, radius=4):
    window = values[max(0, index-radius):min(len(values), index+radius+1)]
    for col in range(values.shape[1]):
        q = window[:, col]
        q = q[np.isfinite(q)]
        if len(q) >= 4 and np.ptp(q) > 1e-8 * max(1., float(np.max(np.abs(q)))):
            return True
    return False


def correlate(ref, target, curves, params, ref_markers, target_markers,
              cancel=None, progress=None):
    """Pairwise fixed-endpoint IDTW. Target markers NEVER influence the path."""
    cancel = cancel or threading.Event()
    progress = progress or (lambda fraction, message: None)
    if ref is target or Path(ref.path).resolve() == Path(target.path).resolve():
        raise ValueError('Выберите две разные скважины.')
    if not curves or len(set(curves)) != len(curves):
        raise ValueError('Выберите хотя бы одну общую кривую без повторов.')
    if len(curves) > 16:
        raise ValueError('Выберите не более 16 кривых для одного расчёта.')
    if not all(c in ref.curves and c in target.curves for c in curves):
        raise ValueError('Выбранная кривая отсутствует в одной из скважин.')
    for name in ('step', 'band', 'gap', 'penalty'):
        value = getattr(params, name)
        if not np.isfinite(value) or value < 0:
            raise ValueError(f'{name}: нужно конечное неотрицательное число.')
    if params.step <= 0 or params.gap <= 0 or params.band <= 0:
        raise ValueError('Шаг, ширина полосы и допустимый разрыв должны быть > 0.')
    if not (50 <= params.max_points <= 4000) or int(params.max_points) != params.max_points:
        raise ValueError('Лимит отсчётов: целое число от 50 до 4000.')
    if not (0 <= params.k <= 10) or int(params.k) != params.k:
        raise ValueError('K: целое число от 0 до 10.')
    if params.normalize not in ('robust', 'none'):
        raise ValueError('Нормировка должна быть robust или none.')
    notes = list(ref.notes) + list(target.notes)
    for curve in curves:
        ru, tu = _key(ref.units.get(curve)), _key(target.units.get(curve))
        if ru != tu:
            notes.append(f'{curve}: единицы различаются ({ru or "?"} / {tu or "?"}); '
                         'автоматического перевода значений ГИС нет.')
            if params.normalize == 'none':
                raise ValueError(f'{curve}: разные единицы ГИС; приведите их к общим единицам '
                                 'или используйте robust для сравнения формы.')
    zr = _grid(ref, params.ref_start, params.ref_end, params.step, params.max_points)
    zt = _grid(target, params.target_start, params.target_end, params.step, params.max_points)
    xr, xt = [], []
    for curve in curves:
        try:
            xr.append(normalize_curve(resample_masked(ref.depth, ref.curves[curve], zr, params.gap), params.normalize))
            xt.append(normalize_curve(resample_masked(target.depth, target.curves[curve], zt, params.gap), params.normalize))
        except ValueError as error:
            raise ValueError(f'{curve}: {error}') from error
    xr, xt = np.column_stack(xr), np.column_stack(xt)
    a = np.column_stack([analytic_masked(xr[:, c]) for c in range(len(curves))])
    b = np.column_stack([analytic_masked(xt[:, c]) for c in range(len(curves))])
    n, m = len(zr), len(zt)
    dzr, dzt = float(zr[1]-zr[0]), float(zt[1]-zt[0])
    width = max(1, int(math.ceil(params.band / dzt)))
    centers = np.linspace(0, m-1, n)
    lower = np.maximum(0, np.floor(centers - width).astype(int))
    upper = np.minimum(m-1, np.ceil(centers + width).astype(int))
    band_size = int(np.max(upper-lower+1))
    if int(np.sum(upper-lower+1)) > 6_000_000:
        raise ValueError('Более 6 млн ячеек DTW: уменьшите полосу или лимит отсчётов.')
    pointers = np.full((n, band_size), 255, dtype=np.uint8)
    previous = np.full(m, np.inf)
    gap_cells = 0
    for i in range(n):
        if cancel.is_set():
            raise Cancelled('Расчёт отменён.')
        js = np.arange(lower[i], upper[i]+1)
        costs, valid = _cost_row(a, b, i, js, params.k)
        gap_cells += int((~valid).sum())
        current = np.full(m, np.inf)
        for q, j in enumerate(js):
            if i == 0 and j == 0:
                current[j] = costs[q]
                pointers[i, q] = 3
                continue
            diagonal = previous[j-1] if i and j else np.inf
            above = previous[j] + params.penalty if i else np.inf
            left = current[j-1] + params.penalty if j else np.inf
            # Prefer diagonal on ties, for a reproducible minimum deformation.
            if diagonal <= above and diagonal <= left:
                best, direction = diagonal, 0
            elif above <= left:
                best, direction = above, 1
            else:
                best, direction = left, 2
            if np.isfinite(best):
                current[j] = costs[q] + best
                pointers[i, q] = direction
        previous = current
        if i % 20 == 0:
            progress(i/n, f'IDTW: {i+1} / {n} строк')
    if not np.isfinite(previous[-1]):
        raise ValueError('В выбранной полосе нет связного пути. Увеличьте ширину полосы.')
    total_cost = float(previous[-1])
    trace, i, j = [], n-1, m-1
    while True:
        trace.append((i, j))
        direction = pointers[i, j-lower[i]]
        if direction == 3:
            break
        if direction == 0:
            i, j = i-1, j-1
        elif direction == 1:
            i -= 1
        elif direction == 2:
            j -= 1
        else:
            raise RuntimeError('Нарушена трассировка пути DTW.')
    path = np.asarray(trace[::-1], dtype=int)
    mapped, similarity = np.full(n, np.nan), np.full(n, np.nan)
    valid_path, boundary_hits, longest, run, last_direction = 0, 0, 0, 0, None
    offset = 0
    for i in range(n):
        if cancel.is_set():
            raise Cancelled('Расчёт отменён.')
        end = offset
        while end < len(path) and path[end, 0] == i:
            end += 1
        js = path[offset:end, 1]
        mapped[i] = float(np.median(zt[js]))
        costs, valid = _cost_row(a, b, i, js, params.k)
        valid_path += int(valid.sum())
        if valid.any():
            similarity[i] = float(np.mean(1-costs[valid]))
        boundary_hits += int(np.sum((np.abs(js-centers[i]) >= width-1) & (js > 0) & (js < m-1)))
        offset = end
    for di, dj in np.diff(path, axis=0):
        direction = (int(di), int(dj))
        run = run + 1 if direction == last_direction and direction != (1, 1) else (0 if direction == (1, 1) else 1)
        longest = max(longest, run)
        last_direction = direction
    valid_fraction = valid_path / len(path)
    if valid_fraction < .25:
        raise ValueError('Менее 25% пути обеспечено данными. Выберите другие интервалы/кривые.')
    if valid_fraction < .9:
        notes.append(f'Данные поддерживают {valid_fraction:.1%} пути; участки пропусков проходят со штрафом 1.25.')
    if boundary_hits / len(path) > .05:
        notes.append('Путь часто касается границы полосы; проверьте ширину полосы и интервалы.')
    if longest > 10:
        notes.append(f'Есть {longest} последовательных недиагональных шагов: проверьте локальное растяжение.')
    if dzr > params.step*1.01 or dzt > params.step*1.01:
        notes.append('Лимит отсчётов увеличил фактический шаг; он указан в статистике.')
    if params.normalize == 'robust':
        notes.append('Применена независимая robust-нормировка кривых (медиана/IQR) до IDTW.')
    if params.k:
        notes.append('Окно semblance обрезает поддержку у концов интервалов; крайние маркеры отклоняются.')
    target_by_name = {}
    for marker in target_markers:
        if _key(marker.name) in target_by_name:
            raise ValueError('Повторяющееся имя маркера в выбранной целевой группе.')
        target_by_name[_key(marker.name)] = marker
    rows = []
    for marker in ref_markers:
        actual = target_by_name.get(_key(marker.name))
        actual_md = actual.md if actual else None
        row = dict(marker=marker.name, ref_md=marker.md, predicted_md=None,
                   actual_md=actual_md, error=None, status='', semblance=None)
        if not zr[0] <= marker.md <= zr[-1]:
            row['status'] = 'Вне опорного интервала'
            rows.append(row)
            continue
        ri = int(np.argmin(abs(zr-marker.md)))
        prediction = float(np.interp(marker.md, zr, mapped))
        tj = int(np.argmin(abs(zt-prediction)))
        local_cost, local_valid = _cost_row(a, b, ri, np.array([tj]), params.k)
        # Require valid support on both sides of interpolation, not just nearest.
        rr = np.clip(np.searchsorted(zr, marker.md), 1, n-1)
        tt = np.clip(np.searchsorted(zt, prediction), 1, m-1)
        common_support = (np.all(np.isfinite(a[rr-1:rr+1]), axis=0) &
                          np.all(np.isfinite(b[tt-1:tt+1]), axis=0))
        common_support &= (np.all(np.isfinite(a[max(0, ri-params.k):ri+params.k+1]), axis=0) &
                           np.all(np.isfinite(b[max(0, tj-params.k):tj+params.k+1]), axis=0))
        # Verify original support as well: a marker can lie inside a very
        # narrow gap that is smaller than the calculation step.
        source_ok = [np.isfinite(resample_masked(ref.depth, ref.curves[c], np.array([marker.md]), params.gap)[0])
                     and np.isfinite(resample_masked(target.depth, target.curves[c], np.array([prediction]), params.gap)[0])
                     for c in curves]
        common_support &= np.array(source_ok)
        supported = np.any(common_support)
        if not supported or not local_valid[0]:
            row['status'] = 'Нет данных для переноса'
        elif not any(common_support[c] and _informative(xr[:, c:c+1], ri)
                     and _informative(xt[:, c:c+1], tj) for c in range(len(curves))):
            row['status'] = 'Мало вариации: неоднозначно'
        else:
            row['predicted_md'] = prediction
            row['semblance'] = float(1-local_cost[0])
            row['status'] = 'Проверить: низкое сходство' if row['semblance'] < .6 else 'Перенесён'
            if actual_md is not None:
                if zt[0] <= actual_md <= zt[-1]:
                    row['error'] = prediction-actual_md
                else:
                    row['status'] += '; контроль вне интервала'
        rows.append(row)
    errors = [row['error'] for row in rows if row['error'] is not None]
    eligible = sum(zr[0] <= m.md <= zr[-1] for m in ref_markers)
    predicted = sum(r['predicted_md'] is not None for r in rows)
    actual_in_range = sum(zt[0] <= m.md <= zt[-1] for m in target_markers)
    stats = {
        'Средний semblance': float(np.nanmean(similarity)),
        'MAE маркеров, м': float(np.mean(np.abs(errors))) if errors else None,
        'RMSE маркеров, м': float(np.sqrt(np.mean(np.square(errors)))) if errors else None,
        'Смещение прогноза, м': float(np.mean(errors)) if errors else None,
        'Макс. ошибка маркера, м': float(np.max(np.abs(errors))) if errors else None,
        'Проверено маркеров': len(errors),
        'Перенесено маркеров': predicted,
        'Маркеров в опорном интервале': eligible,
        'Покрытие маркеров': predicted/eligible if eligible else None,
        'Контрольных маркеров в целевом интервале': actual_in_range,
        'Доля проверенных целевых маркеров': len(errors)/actual_in_range if actual_in_range else None,
        'Доля пути с данными': valid_fraction,
        'Шаг опорной сетки, м': dzr,
        'Шаг целевой сетки, м': dzt,
        'Число ячеек полосы': int(np.sum(upper-lower+1)),
        'Стоимость пути со штрафами': total_cost,
        'Макс. серия недиагональных шагов': longest,
        'Нормировка': params.normalize,
        'Предупреждения': '\n'.join(notes) or 'Нет',
    }
    progress(1., 'Готово')
    return Result(ref, target, curves, params, zr, zt, xr, xt, path, mapped,
                  similarity, rows, stats, notes, list(ref_markers), list(target_markers))


def _synthetic_case():
    """Known nonlinear depth map, independent of the fitting algorithm."""
    zr, zt = np.linspace(0, 100, 401), np.linspace(0, 120, 481)
    def mapping(z):
        return 1.2*z + 6*np.sin(np.pi*z/100)
    peaks = [(9,2,1.2), (22,3,-.8), (37,1.5,1.5), (56,4,-1.3), (76,2,1), (92,1.7,-.7)]
    def signal(z):
        return sum(amplitude*np.exp(-.5*((z-center)/width)**2)
                   for center, width, amplitude in peaks) + .08*np.sin(.9*z)
    inverse = np.interp(zt, mapping(zr), zr)
    ref = Well('synthetic_reference.las', 'Синтетическая А', '0001', zr,
               {'GR': signal(zr)}, {'GR': 'API'})
    target = Well('synthetic_target.las', 'Синтетическая Б', '0002', zt,
                  {'GR': signal(inverse)}, {'GR': 'API'})
    rm = [Marker(f'M{i+1}', ref.name, ref.uwi, float(p[0])) for i,p in enumerate(peaks)]
    tm = [Marker(m.name, target.name, target.uwi, float(mapping(m.md))) for m in rm]
    return ref, target, rm, tm


def _fixture_las(well, wrap=False, feet=False, version='2.0'):
    factor = 1/.3048 if feet else 1
    z = well.depth*factor
    unit = 'FT' if feet else 'M'
    header = (f'~Version Information\nVERS. {version} : VERSION\n'
              f'WRAP. {"YES" if wrap else "NO"} : WRAPPED\n'
              f'~Well Information\nSTRT.{unit} {z[0]:.9f} : START\n'
              f'STOP.{unit} {z[-1]:.9f} : STOP\nSTEP.{unit} {z[1]-z[0]:.9f} : STEP\n'
              f'NULL. -999.25 : NULL\nWELL. {well.name} : WELL\nUWI. {well.uwi} : ID\n'
              f'~Curve Information\nDEPT.{unit} : DEPTH\nGR.API : GR\n~ASCII\n')
    lines = []
    for depth, value in zip(z, well.curves['GR']):
        value = value if np.isfinite(value) else -999.25
        lines.append(f'{depth:.9f}' + ('\n' if wrap else ' ') + f'{value:.9f}')
    return header + '\n'.join(lines) + '\n'


def self_test():
    """Run real numerical and input-format regressions without opening Tk."""
    checked = []
    def check(condition, message):
        if not condition:
            raise AssertionError(message)
        checked.append(message)
    a = np.array([1+2j, 2-1j, -.4+.8j])[:, None]
    for gain, expected in [(1,0), (2,.1), (-1,1), (0,.5)]:
        cost, valid = _cost_row(a, a*gain, 1, np.array([1]), 1)
        check(valid[0] and abs(cost[0]-expected) < 1e-12, f'cost gain={gain}')
    zero = np.zeros_like(a)
    cost, valid = _cost_row(zero, zero, 1, np.array([1]), 0)
    check(not valid[0], 'zero energy is invalid')
    n = 128
    phase = 2*np.pi*5*np.arange(n)/n
    check(np.max(np.abs(analytic_signal(np.cos(phase))-np.exp(1j*phase))) < 1e-12,
          'FFT Hilbert sign and amplitude')
    ref, target, rm, tm = _synthetic_case()
    p = Params(step=.25, band=15, penalty=.025)
    result = correlate(ref, target, ['GR'], p, rm, tm)
    mae = result.stats['MAE маркеров, м']
    affine_mae = np.mean([abs(1.2*r.md-t.md) for r,t in zip(rm,tm)])
    check(mae < .5 and mae < affine_mae/5, 'nonlinear withheld marker accuracy')
    check(np.all(np.diff(result.mapped) >= 0), 'monotone depth mapping')
    without = correlate(ref, target, ['GR'], p, rm, [])
    check(np.array_equal(result.path, without.path), 'target markers do not affect path')
    # A source gap smaller than the resampling step must not disappear.
    z = np.round(np.linspace(0,100,1001), 4)
    values = np.sin(.23*z)+.3*np.sin(.81*z)
    missing = values.copy()
    missing[522] = np.nan
    r_gap = Well('gap.las', 'Gap', '', z, {'GR': missing}, {'GR':'API'})
    t_gap = Well('complete.las', 'Complete', '', z, {'GR': values}, {'GR':'API'})
    gap_result = correlate(r_gap, t_gap, ['GR'], Params(step=.5, band=5),
                           [Marker('NULL', 'Gap', '', 52.2), Marker('OUT','Gap','',110)], [])
    check(all(row['predicted_md'] is None for row in gap_result.rows), 'gap and out-of-range rejection')
    check(np.isnan(resample_masked(z, missing, np.arange(0,100.5,.5),3)).any(),
          'coarse grid preserves narrow acquisition gap')
    flat = Well('flat.las','Flat','',ref.depth,{'GR':np.ones(len(ref.depth))},{'GR':'API'})
    try:
        correlate(flat, target, ['GR'], p, [], [])
    except ValueError:
        checked.append('flat curve rejected')
    else:
        raise AssertionError('flat curve was accepted')
    event = threading.Event()
    event.set()
    try:
        correlate(ref, target, ['GR'], p, [], [], cancel=event)
    except Cancelled:
        checked.append('cancellation')
    else:
        raise AssertionError('cancellation ignored')
    # Actual LAS and XLSX round trips are part of the integration test.
    from openpyxl import Workbook
    with tempfile.TemporaryDirectory(prefix='idtw_checks_') as directory:
        directory = Path(directory)
        for wrap, feet, encoding in [(False,False,'utf-8'),(True,False,'cp1251'),(False,True,'utf-8')]:
            laspath = directory/'well.las'
            laspath.write_text(_fixture_las(ref,wrap,feet),encoding=encoding)
            loaded = load_las(laspath)
            check(loaded.name == ref.name and loaded.uwi == ref.uwi and
                  np.allclose(loaded.depth,ref.depth,atol=1e-6) and
                  np.allclose(loaded.curves['GR'],ref.curves['GR'],atol=1e-6),
                  f'LAS wrap={wrap} feet={feet} encoding={encoding}')
        laspath.write_text(_fixture_las(ref,version='99.0'),encoding='utf-8')
        try:
            load_las(laspath)
        except ValueError:
            checked.append('invalid LAS version rejected')
        else:
            raise AssertionError('invalid LAS version accepted')
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Маркер','Скважина','UWI','MD','X','Y','Z'])
        for marker in rm+tm:
            sheet.append([marker.name,marker.well,marker.uwi,marker.md,None,None,None])
        xlsx = directory/'markers.xlsx'
        workbook.save(xlsx)
        workbook.close()
        loaded = load_markers(xlsx)
        groups = marker_groups(loaded)
        check(len(loaded)==12 and len(groups)==2 and bool(match_group(ref,groups)), 'XLSX and UWI matching')
        check(loaded[0].uwi=='0001', 'UWI leading zeros')
    print(f'PASS: {len(checked)} checks; nonlinear marker MAE={mae:.4f} m; '
          f'affine baseline MAE={affine_mae:.4f} m.')
    for name in checked:
        print('  OK:', name)
    return result


@dataclass
class RunBundle:
    result: Result
    candidates: list = field(default_factory=list)
    report: list = field(default_factory=list)
    config: dict = field(default_factory=dict)


def clone_result(result):
    # Wells are immutable input snapshots. Avoid copying full LAS data once per
    # ensemble result, and keep snapshot identity when projects are saved.
    clone=copy.deepcopy(result,{id(result.ref):result.ref,id(result.target):result.target})
    return clone


def auto_intervals(ref, target, markers, params, ref_margin=10., target_margin=100., shift=0.):
    """Only reference markers determine interval selection; never target labels."""
    if not markers:
        raise ValueError('Для автоматических интервалов нужны маркеры опорной скважины.')
    if not all(np.isfinite(v) for v in (ref_margin, target_margin, shift)) or min(ref_margin, target_margin) < 0:
        raise ValueError('Запасы должны быть конечными неотрицательными числами.')
    low, high = min(m.md for m in markers), max(m.md for m in markers)
    rs, re = max(float(ref.depth[0]), low-ref_margin), min(float(ref.depth[-1]), high+ref_margin)
    ts, te = max(float(target.depth[0]), low+shift-target_margin), min(float(target.depth[-1]), high+shift+target_margin)
    if rs >= re or ts >= te:
        raise ValueError('Интервал по маркерам не пересекается с LAS. Проверьте смещение целевой скважины.')
    return replace(params, ref_start=rs, ref_end=re, target_start=ts, target_end=te)


def candidate_settings(base, curves, count=16, seed=42):
    """Reproducible finite search. Vary shape parameters, never target labels.

    Keep normalisation/intervals fixed to preserve user interpretation. Include
    full-channel runs and leave-one-channel-out checks. No duplicate settings.
    """
    if not 2 <= count <= 64:
        raise ValueError('Число генераций должно быть от 2 до 64.')
    if not curves:
        raise ValueError('Не выбраны кривые.')
    curve_sets = [tuple(curves)]
    if len(curves) > 1:
        curve_sets += [tuple(c for c in curves if c != excluded) for excluded in curves]
    choices = list(itertools.product((.5, 1., 1.5), (.75, 1., 1.25),
                                    sorted({0, 1, 2, base.k}),
                                    sorted({0., .025, .075, .15, base.penalty}), curve_sets))
    rng = np.random.default_rng(seed)
    rng.shuffle(choices)
    result = [(replace(base), list(curves))]
    seen = {(base.band, base.step, base.k, base.penalty, tuple(curves))}
    for band, step, k, penalty, channels in choices:
        p = replace(base, band=base.band*band, step=base.step*step, k=k, penalty=penalty)
        key = (p.band, p.step, p.k, p.penalty, channels)
        if key in seen:
            continue
        result.append((p, list(channels)))
        seen.add(key)
        if len(result) == count:
            break
    return result


def _quality(result):
    """Heuristic diagnostic, not probability. Never uses actual/error columns."""
    support = float(result.stats['Доля пути с данными'])
    coverage = result.stats.get('Покрытие маркеров')
    coverage = 1. if coverage is None else float(coverage)
    profile = np.diff(result.mapped) / np.diff(result.zr)
    expected = (result.zt[-1]-result.zt[0])/(result.zr[-1]-result.zr[0])
    deformation = float(np.mean((profile < expected*.2) | (profile > expected*5)))
    center = result.zt[0] + (result.zr-result.zr[0])*expected
    boundary = float(np.mean(abs(result.mapped-center) >= result.params.band*.9))
    similarity = float(np.nanmean(result.similarity))
    score = .40*similarity + .30*support + .30*coverage - .15*deformation - .10*boundary
    return dict(quality=score, similarity=similarity, support=support, coverage=coverage,
                deformation=deformation, boundary=boundary)


def build_consensus(candidates, report, tolerance, requested):
    """Cluster entire depth mappings; choose a real medoid, never splice tops.

    Complete-link clusters prevent chaining disparate modes through intermediate
    solutions. RMS depth distance is measured on reference marker positions plus
    a sparse background grid. Display spread per marker within the selected mode.
    """
    if not candidates:
        raise ValueError('Ни одна генерация не завершилась. Проверьте интервалы и данные.')
    if tolerance <= 0 or not np.isfinite(tolerance):
        raise ValueError('Допуск консенсуса должен быть положительным.')
    first = candidates[0]
    landmarks = [m.md for m in first.reference_markers if first.zr[0] <= m.md <= first.zr[-1]]
    sample = np.unique(np.r_[np.linspace(first.zr[0], first.zr[-1], 17), landmarks])
    vectors = np.array([np.interp(sample, r.zr, r.mapped) for r in candidates])
    distances = np.sqrt(np.mean((vectors[:, None, :]-vectors[None, :, :])**2, axis=2))
    qualities = [_quality(r) for r in candidates]
    clusters = []
    for index in sorted(range(len(candidates)), key=lambda i: -qualities[i]['quality']):
        fitting = [g for g in clusters if all(distances[index,j] <= tolerance for j in g)]
        if fitting:
            min(fitting, key=lambda g: np.mean(distances[index,g])).append(index)
        else:
            clusters.append([index])
    clusters.sort(key=lambda g: (-len(g), -np.mean([qualities[i]['quality'] for i in g])))
    dominant = clusters[0]
    chosen = min(dominant, key=lambda i: (float(np.mean(distances[i, dominant])), -qualities[i]['quality']))
    result = clone_result(candidates[chosen])
    credible_competition = len(clusters)>1 and len(clusters[1]) >= max(2, .5*len(dominant))
    by_name = [{_key(row['marker']):row for row in r.rows} for r in candidates]
    for row in result.rows:
        name = _key(row['marker'])
        values = [by_name[i][name]['predicted_md'] for i in dominant
                  if name in by_name[i] and by_name[i][name]['predicted_md'] is not None]
        row['support'] = len(values)/requested
        row['p10'], row['p90'] = (map(float,np.percentile(values,[10,90])) if values else (None,None))
        row['alternatives'] = []
        for group in clusters[1:]:
            alt = [by_name[i][name]['predicted_md'] for i in group
                   if name in by_name[i] and by_name[i][name]['predicted_md'] is not None]
            if alt:
                row['alternatives'].append({'md':float(np.median(alt)), 'votes':len(alt), 'total':requested})
        reasons = []
        if row['predicted_md'] is not None:
            if len(values)<3:
                reasons.append('мало реализаций')
            if row['support'] < .6:
                reasons.append('слабая поддержка генераций')
            if values and row['p90']-row['p10'] > 2*tolerance:
                reasons.append('широкий диапазон')
            if credible_competition:
                reasons.append('конкурирующее решение')
            if reasons:
                row['status'] = 'Проверить: ' + '; '.join(reasons)
        row['confirmed'] = False
    for idx, quality in enumerate(qualities):
        generation = candidates[idx].provenance['generation']
        item = next(item for item in report if item['generation']==generation)
        item.update(quality)
        item['cluster'] = next(k+1 for k,g in enumerate(clusters) if idx in g)
        item['selected'] = idx == chosen
    result.provenance.update(mode='consensus', selected_generation=candidates[chosen].provenance['generation'],
                             requested=requested, successful=len(candidates), tolerance=tolerance)
    result.stats.update({'Генераций успешно / запрошено': f'{len(candidates)} / {requested}',
                         'Групп решений':len(clusters), 'Поддержка группы консенсуса':len(dominant)/requested,
                         'Выбранная генерация':candidates[chosen].provenance['generation'],
                         'Диапазон P10–P90':'Чувствительность внутри выбранной группы; не доверительный интервал'})
    return result


def run_ensemble(ref, target, curves, params, rm, tm, count=16, seed=42,
                 tolerance=2., cancel=None, progress=None, tuning=False):
    cancel = cancel or threading.Event()
    progress = progress or (lambda f,s:None)
    options = candidate_settings(params, curves, count, seed)
    if tuning:
        rnames = {_key(m.name) for m in rm if (params.ref_start if params.ref_start is not None else ref.depth[0]) <= m.md <=
                  (params.ref_end if params.ref_end is not None else ref.depth[-1])}
        controls = { _key(m.name):m for m in tm if _key(m.name) in rnames and
                    (params.target_start if params.target_start is not None else target.depth[0]) <= m.md <=
                    (params.target_end if params.target_end is not None else target.depth[-1]) }
        if len(controls) < 2:
            raise ValueError('Для подбора нужны минимум два общих контрольных маркера внутри интервалов.')
    candidates, report = [], []
    for idx, (p, channels) in enumerate(options):
        if cancel.is_set():
            raise Cancelled('Расчёт отменён.')
        item = dict(generation=idx+1, params=asdict(p), curves=channels, selected=False)
        try:
            result = correlate(ref,target,channels,p,rm,tm,cancel,
                               lambda f,s:progress((idx+f)/len(options),f'Генерация {idx+1}/{len(options)}: {s}'))
            result.provenance = dict(generation=idx+1,mode='candidate',seed=seed)
            item.update(_quality(result), status='OK')
            if tuning:
                errors = { _key(row['marker']):abs(row['error']) for row in result.rows
                           if row['error'] is not None and _key(row['marker']) in controls }
                missing_penalty = (result.zt[-1]-result.zt[0]) + tolerance
                loss = np.mean([errors.get(name,missing_penalty) for name in controls])
                item.update(loss=float(loss), control_count=len(controls), checked=len(errors),
                            control_coverage=len(errors)/len(controls),
                            mae=float(np.mean(list(errors.values()))) if errors else None,
                            within_tolerance=sum(e<=tolerance for e in errors.values())/len(controls))
            candidates.append(result)
        except Cancelled:
            raise
        except ValueError as error:
            item.update(status=str(error),quality=None)
        report.append(item)
    if not candidates:
        raise ValueError('Все генерации отклонены: ' + '; '.join(item['status'] for item in report[:3]))
    if tuning:
        successful = [x for x in report if x['status']=='OK' and x['checked']>=2]
        if not successful:
            raise ValueError('Ни один вариант не предсказал хотя бы два контрольных маркера. Подбор ненадёжен: проверьте покрытие.')
        best = min(successful, key=lambda x:(x['loss'], -x['control_coverage'], -x['quality']))
        best['selected'] = True
        result = clone_result(next(r for r in candidates if r.provenance['generation']==best['generation']))
        result.provenance.update(mode='tuned',control_names=sorted(controls),search_count=len(options))
        result.stats.update({'Режим оценки':'Подбор на этих контрольных маркерах; НЕ независимая проверка',
                             'Потеря со штрафом за пропуски, м':best['loss'],
                             'Покрытие контрольных маркеров':best['control_coverage'],
                             'Доля ошибок в допуске':best['within_tolerance'],
                             'Выбрана генерация':best['generation']})
    else:
        result = build_consensus(candidates, report, tolerance, len(options))
    return RunBundle(result,candidates,report,dict(ensemble=True,count=count,seed=seed,tolerance=tolerance,tuning=tuning))


def output_markers(result):
    return [Marker(row['marker'], result.target.name,result.target.uwi,float(row['predicted_md']))
            for row in result.rows if row['predicted_md'] is not None]


def run_sequence(wells, groups_by_path, curves, params, config, cancel=None, progress=None, initial_markers=None, manual_picks=None):
    cancel = cancel or threading.Event()
    progress = progress or (lambda f,s:None)
    if len(wells)<2 or len({w.path for w in wells})!=len(wells):
        raise ValueError('Нужны минимум две разные скважины в наборе.')
    rm = list(initial_markers if initial_markers is not None else groups_by_path.get(wells[0].path,[]))
    if not rm:
        raise ValueError('У первой скважины набора нет опорных маркеров.')
    bundles, errors, inherited = [], [], set()
    for idx,(ref,target) in enumerate(zip(wells,wells[1:])):
        if cancel.is_set():
            raise Cancelled('Расчёт отменён.')
        channels = [c for c in curves if c in ref.curves and c in target.curves]
        try:
            if not channels:
                raise ValueError('Нет выбранных общих кривых.')
            p = auto_intervals(ref,target,rm,params,config['ref_margin'],config['target_margin'],config['shift']) if config['auto'] else replace(params)
            tm = groups_by_path.get(target.path,[])
            cb = lambda f,s:progress((idx+f)/(len(wells)-1),f'{ref.name} → {target.name}: {s}')
            if config['ensemble']:
                bundle = run_ensemble(ref,target,channels,p,rm,tm,config['count'],config['seed'],config['tolerance'],cancel,cb)
            else:
                bundle = RunBundle(correlate(ref,target,channels,p,rm,tm,cancel,cb))
            bundle.config.update(config)
            bundle.config['requested_curves']=list(curves)
            bundle.config['base_params']=asdict(params)
            reapply_manual(bundle.result,(manual_picks or {}).get(target.path,{}))
            for row in bundle.result.rows:
                if _key(row['marker']) in inherited:
                    row['status'] += '; проверить: неопределённость предыдущего переноса'
                    row['inherited_uncertainty'] = True
            bundles.append(bundle)
            inherited |= {_key(row['marker']) for row in bundle.result.rows
                          if 'провер' in row['status'].lower() and row['predicted_md'] is not None}
            rm = output_markers(bundle.result)
            if not rm and idx<len(wells)-2:
                raise ValueError('Нет перенесённых маркеров для следующей пары.')
        except ValueError as error:
            errors.append(f'{ref.name} → {target.name}: {error}')
            break
    return bundles, errors


def edit_marker(result, name, md, comment='', validate_order=True):
    """Manual pick overrides exported tops, not the fitted log-to-log path."""
    if not np.isfinite(md) or not result.zt[0]<=md<=result.zt[-1]:
        raise ValueError('MD должна лежать внутри целевого расчётного интервала.')
    row = next((r for r in result.rows if r['marker']==name),None)
    if row is None:
        raise ValueError('Маркер не найден.')
    for other in result.rows if validate_order else []:
        if other is row or other['predicted_md'] is None:
            continue
        if ((other['ref_md']<row['ref_md'] and other['predicted_md']>=md) or
            (other['ref_md']>row['ref_md'] and other['predicted_md']<=md)):
            raise ValueError('Исправление нарушает порядок маркеров. Проверьте соседние границы.')
    row.setdefault('original_md',row['predicted_md'])
    row.setdefault('history',[]).append(dict(previous_md=row['predicted_md'],new_md=float(md),
                                          comment=comment,time=time.strftime('%Y-%m-%d %H:%M:%S')))
    row.update(predicted_md=float(md),confirmed=True,status='Подтверждён пользователем',
               error=float(md-row['actual_md']) if row['actual_md'] is not None else None,
               semblance=None)
    refresh_marker_stats(result)


def refresh_marker_stats(result):
    # Do not report manually corrected rows as successful automatic predictions.
    errors = [r['error'] for r in result.rows if r.get('error') is not None and not r.get('confirmed')]
    control_count=result.stats.get('Контрольных маркеров в целевом интервале',0)
    predicted=sum(r['predicted_md'] is not None and result.zr[0]<=r['ref_md']<=result.zr[-1] for r in result.rows)
    eligible=result.stats.get('Маркеров в опорном интервале',0)
    result.stats.update({'MAE маркеров, м':float(np.mean(np.abs(errors))) if errors else None,
                         'RMSE маркеров, м':float(np.sqrt(np.mean(np.square(errors)))) if errors else None,
                         'Смещение прогноза, м':float(np.mean(errors)) if errors else None,
                         'Макс. ошибка маркера, м':float(np.max(np.abs(errors))) if errors else None,
                         'Проверено маркеров':len(errors),
                         'Доля проверенных целевых маркеров':len(errors)/control_count if control_count else None,
                         'Перенесено маркеров':predicted,
                         'Покрытие маркеров':predicted/eligible if eligible else None,
                         'Подтверждено пользователем':sum(bool(r.get('confirmed')) for r in result.rows)})


def reapply_manual(result, picks):
    """Keep user picks across recalculation; never silently replace them."""
    proposed={_key(row['marker']):picks.get(_key(row['marker']),row).get('predicted_md') for row in result.rows}
    for row in result.rows:
        saved=picks.get(_key(row['marker']))
        if not saved:
            continue
        try:
            md=saved['predicted_md']
            for other in result.rows:
                other_md=proposed[_key(other['marker'])]
                if other is row or other_md is None:
                    continue
                if ((other['ref_md']<row['ref_md'] and other_md>=md) or
                    (other['ref_md']>row['ref_md'] and other_md<=md)):
                    raise ValueError('Подтверждённая глубина конфликтует с порядком других границ.')
            edit_marker(result,row['marker'],md,'Повторное применение подтверждённой границы',validate_order=False)
            row['history']=copy.deepcopy(saved.get('history',[]))
            row['original_md']=saved.get('original_md')
        except ValueError as error:
            row.update(predicted_md=None,error=None,status='Проверить ручную границу: '+str(error))
            row['history']=copy.deepcopy(saved.get('history',[]))
            row['saved_manual_md']=saved['predicted_md']
    if picks:
        refresh_marker_stats(result)


def marker_color(name):
    hue = int(hashlib.sha256(_key(name).encode('utf-8')).hexdigest()[:8],16)/0xffffffff
    rgb = colorsys.hsv_to_rgb(hue,.75,.68)
    return '#'+''.join(f'{round(c*255):02x}' for c in rgb)


def flag_reused_calibration(bundles, contexts):
    for bundle in bundles:
        context=contexts.get(bundle.result.target.path)
        if context:
            bundle.result.provenance['calibrated_on_target']=copy.deepcopy(context)
            bundle.result.stats['Режим оценки']='Параметры ранее подбирались по этой целевой скважине; НЕ независимая проверка'


def _json_clean(value):
    if isinstance(value, dict):
        return {str(k):_json_clean(v) for k,v in value.items()}
    if isinstance(value,(list,tuple)):
        return [_json_clean(v) for v in value]
    if isinstance(value,np.generic):
        value = value.item()
    if isinstance(value,float) and not math.isfinite(value):
        return None
    return value


def save_project(path, wells, markers, bundles, ui):
    """Atomic self-contained zip/JSON/NumPy snapshot; no pickle, no sidecars."""
    from io import BytesIO
    path = Path(path)
    fd, temporary_name = tempfile.mkstemp(prefix=path.name+'.',suffix='.tmp',dir=path.parent)
    os.close(fd)
    temporary = Path(temporary_name)
    snapshots=list(wells)
    well_ids = {id(well):str(i) for i,well in enumerate(snapshots)}
    active=[well_ids[id(well)] for well in wells]
    for bundle in bundles:
        for result in [bundle.result]+bundle.candidates:
            for well in (result.ref,result.target):
                if id(well) not in well_ids:
                    well_ids[id(well)]=str(len(snapshots))
                    snapshots.append(well)
    manifest = dict(format='idtw-project',version=2,wells=[],active_well_ids=active,markers=[asdict(m) for m in markers],bundles=[],ui=ui)
    try:
        with zipfile.ZipFile(temporary,'w',compression=zipfile.ZIP_DEFLATED) as archive:
            counter = itertools.count()
            total_bytes=0
            def array(value):
                nonlocal total_bytes
                name = f'arrays/{next(counter)}.npy'
                stream = BytesIO()
                np.save(stream,value,allow_pickle=False)
                total_bytes+=stream.tell()
                if total_bytes>500*1024*1024:
                    raise ValueError('Проект превышает 500 МБ массивов. Уменьшите число генераций или скважин.')
                archive.writestr(name,stream.getvalue())
                return name
            def encode_result(r):
                return dict(ref=well_ids[id(r.ref)],target=well_ids[id(r.target)],curves=r.curves,params=asdict(r.params),
                            arrays={k:array(getattr(r,k)) for k in ('zr','zt','xr','xt','path','mapped','similarity')},
                            rows=r.rows,stats=r.stats,notes=r.notes,reference_markers=[asdict(m) for m in r.reference_markers],
                            control_markers=[asdict(m) for m in r.control_markers],provenance=r.provenance)
            for well in snapshots:
                manifest['wells'].append(dict(id=well_ids[id(well)],path=well.path,name=well.name,uwi=well.uwi,
                                             depth=array(well.depth),curves={c:array(v) for c,v in well.curves.items()},
                                             units=well.units,notes=well.notes))
            for bundle in bundles:
                manifest['bundles'].append(dict(result=encode_result(bundle.result),
                                               candidates=[encode_result(r) for r in bundle.candidates],
                                               report=bundle.report,config=bundle.config))
            archive.writestr('project.json',json.dumps(_json_clean(manifest),ensure_ascii=False,allow_nan=False))
        os.replace(temporary,path)
    finally:
        if temporary.exists():
            temporary.unlink()


def load_project(path):
    from io import BytesIO
    with zipfile.ZipFile(path) as archive:
        if sum(i.file_size for i in archive.infolist())>512*1024*1024:
            raise ValueError('Распакованный проект превышает 512 МБ.')
        doc = json.loads(archive.read('project.json'))
        if doc.get('format')!='idtw-project' or doc.get('version')!=2:
            raise ValueError('Неизвестный формат или версия проекта.')
        def array(name):
            value = np.load(BytesIO(archive.read(name)),allow_pickle=False)
            if value.dtype.kind not in 'fiu' or value.ndim>2:
                raise ValueError('Неверный массив в проекте.')
            return value
        wells, lookup = [], {}
        for spec in doc['wells']:
            well = Well(spec['path'],spec['name'],spec['uwi'],array(spec['depth']),
                        {c:array(v) for c,v in spec['curves'].items()},spec['units'],spec['notes'])
            if well.depth.ndim!=1 or len(well.depth)<8 or not np.all(np.isfinite(well.depth)) or np.any(np.diff(well.depth)<=0):
                raise ValueError('Некорректная глубина в проекте.')
            if any(v.shape!=well.depth.shape for v in well.curves.values()):
                raise ValueError('Размеры кривых не совпадают с глубиной.')
            lookup[spec['id']] = well
            wells.append(well)
        def decode_result(spec):
            a = {k:array(v) for k,v in spec['arrays'].items()}
            if a['mapped'].shape != a['zr'].shape or a['similarity'].shape!=a['zr'].shape:
                raise ValueError('Повреждённое соответствие глубин в проекте.')
            return Result(lookup[spec['ref']],lookup[spec['target']],spec['curves'],Params(**spec['params']),
                          **a,rows=spec['rows'],stats=spec['stats'],notes=spec['notes'],
                          reference_markers=[Marker(**m) for m in spec['reference_markers']],
                          control_markers=[Marker(**m) for m in spec['control_markers']],provenance=spec['provenance'])
        bundles = [RunBundle(decode_result(b['result']),[decode_result(r) for r in b['candidates']],b['report'],b['config'])
                   for b in doc['bundles']]
        if 'active_well_ids' in doc:
            wells=[lookup[key] for key in doc['active_well_ids']]
        return wells,[Marker(**m) for m in doc['markers']],bundles,doc['ui']


def advanced_self_test():
    ref,target,rm,tm=_synthetic_case()
    base=Params(step=.5,band=15,penalty=.025)
    checks=[]
    def check(condition,name):
        if not condition:
            raise AssertionError(name)
        checks.append(name)
    auto=auto_intervals(ref,target,rm,base,2,8,20)
    check((auto.ref_start,auto.ref_end,auto.target_start,auto.target_end)==(7.,94.,21.,120.),'auto intervals, margins, shift and clipping')
    options=candidate_settings(base,['GR'],8,123)
    check([(asdict(p),c) for p,c in options]==[(asdict(p),c) for p,c in candidate_settings(base,['GR'],8,123)],'reproducible parameter search')
    check(len({json.dumps(asdict(p),sort_keys=True)+str(c) for p,c in options})==8,'unique candidate settings')
    ensemble=run_ensemble(ref,target,['GR'],base,rm,tm,count=6,seed=3,tolerance=2)
    check(len(ensemble.candidates)==6 and any(item.get('selected') for item in ensemble.report),'ensemble candidates and selected medoid')
    check(any(np.array_equal(ensemble.result.path,r.path) for r in ensemble.candidates),'consensus is a real coherent path')
    check(all(row.get('p10')<=row['p90'] for row in ensemble.result.rows if row.get('p10') is not None),'uncertainty bounds')
    independent=run_ensemble(ref,target,['GR'],base,rm,[],count=6,seed=3,tolerance=2)
    check(ensemble.result.provenance['selected_generation']==independent.result.provenance['selected_generation'] and
          np.array_equal(ensemble.result.path,independent.result.path),'consensus selection does not use control labels')
    tuned=run_ensemble(ref,target,['GR'],base,rm,tm,count=6,seed=3,tolerance=2,tuning=True)
    winner=next(item for item in tuned.report if item.get('selected'))
    check(winner['loss']==min(item['loss'] for item in tuned.report if item['status']=='OK'),'tuning minimises declared loss')
    check(tuned.result.provenance['mode']=='tuned' and 'НЕ независимая' in tuned.result.stats['Режим оценки'],'tuning provenance is explicit')
    # Distinct modes must remain alternatives, not be averaged into a false pick.
    modes=[]
    report=[]
    for i in range(5):
        r=copy.deepcopy(ensemble.candidates[0])
        r.provenance['generation']=i+1
        if i>=3:
            r.mapped=r.mapped+8
            for row in r.rows:
                if row['predicted_md'] is not None:
                    row['predicted_md']+=8
        modes.append(r)
        report.append(dict(generation=i+1,status='OK'))
    voted=build_consensus(modes,report,2,7)
    check(voted.stats['Групп решений']==2 and bool(voted.rows[0]['alternatives']),'competing modes retained')
    check(abs(voted.rows[0]['support']-3/7)<1e-9,'failed generations remain in support denominator')
    edited=copy.deepcopy(ensemble.result)
    original=edited.rows[0]['predicted_md']
    edit_marker(edited,edited.rows[0]['marker'],original+.2,'test interpretation')
    check(edited.rows[0]['confirmed'] and edited.rows[0]['original_md']==original and len(edited.rows[0]['history'])==1,'manual pick history and provenance')
    check(edited.stats['Проверено маркеров']==5,'manual corrections excluded from automatic validation')
    try:
        edit_marker(edited,edited.rows[0]['marker'],edited.rows[1]['predicted_md']+1)
    except ValueError:
        checks.append('manual crossing rejected')
    else:
        raise AssertionError('manual crossing accepted')
    third=Well('synthetic_third.las','Синтетическая В','0003',target.depth+15,
               {'GR':target.curves['GR'].copy()},target.units.copy())
    controls3=[Marker(m.name,third.name,third.uwi,m.md+15) for m in tm]
    config=dict(auto=False,ensemble=False,count=4,seed=42,tolerance=2,ref_margin=10,target_margin=100,shift=0)
    bundles,errors=run_sequence([ref,target,third],{ref.path:rm,target.path:tm,third.path:controls3},['GR'],base,config)
    check(len(bundles)==2 and not errors,'three-well sequence')
    check(np.allclose([m.md for m in bundles[1].result.reference_markers],
                      [r['predicted_md'] for r in bundles[0].result.rows]),'sequence propagates predictions, not target control labels')
    edit_marker(bundles[0].result,bundles[0].result.rows[0]['marker'],bundles[0].result.rows[0]['predicted_md']+.3)
    repeated,errors=run_sequence([target,third],{third.path:controls3},['GR'],base,config,
                                 initial_markers=output_markers(bundles[0].result))
    check(abs(repeated[0].result.reference_markers[0].md-bundles[0].result.rows[0]['predicted_md'])<1e-12,'downstream recalculation uses edited picks')
    with tempfile.TemporaryDirectory(prefix='idtw_project_check_') as directory:
        path=Path(directory)/'roundtrip.idtw'
        ensemble.result=edited
        save_project(path,[ref,target,third],rm+tm+controls3,[ensemble,tuned]+bundles,{'sequence':[ref.path,target.path,third.path]})
        ws,ms,bs,ui=load_project(path)
        check(len(ws)==3 and len(ms)==18 and len(bs)==4 and len(bs[0].candidates)==6,'self-contained project round trip')
        check(bs[0].result.rows[0]['history']==edited.rows[0]['history'] and
              np.array_equal(bs[0].result.path,edited.path) and ui['sequence'][2]==third.path,'project preserves edits, paths and order')
        check(np.allclose(ws[0].curves['GR'],ref.curves['GR']),'project restores data without source files')
        changed=Well(ref.path,ref.name,ref.uwi,ref.depth.copy(),{'GR':ref.curves['GR']+100},ref.units.copy())
        save_project(path,[changed,target],rm+tm,[ensemble],{})
        ws,ms,bs,ui=load_project(path)
        check(np.allclose(ws[0].curves['GR'],ref.curves['GR']+100) and
              np.allclose(bs[0].result.ref.curves['GR'],ref.curves['GR']),
              'project preserves old result snapshots after reloading the same LAS path')
    restored=clone_result(ensemble.candidates[0])
    picks={_key(edited.rows[0]['marker']):edited.rows[0]}
    reapply_manual(restored,picks)
    check(restored.rows[0]['confirmed'] and restored.rows[0]['history']==edited.rows[0]['history'],
          'manual pick survives recalculation without rewriting its history')
    event=threading.Event()
    event.set()
    try:
        run_ensemble(ref,target,['GR'],base,rm,tm,count=4,cancel=event)
    except Cancelled:
        checks.append('ensemble cancellation')
    else:
        raise AssertionError('ensemble cancellation ignored')
    try:
        run_ensemble(ref,target,['GR'],base,rm,tm[:1],count=4,tuning=True)
    except ValueError:
        checks.append('insufficient calibration controls rejected')
    else:
        raise AssertionError('single-control tuning accepted')
    print(f'PASS: {len(checks)} advanced checks.')
    for name in checks:
        print('  OK:',name)
    return len(checks)


import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class PairApp(tk.Tk):
    """Tk is used only by the main thread; workers communicate via a queue."""

    def __init__(self):
        super().__init__()
        self.title("IDTW — корреляция каротажа двух скважин")
        self.geometry("1240x880")
        self.minsize(1020, 730)
        self.wells = {}
        self.markers = []
        self.groups = {}
        self.assignments = {}
        self.result = None
        self.busy = False
        self.cancel_event = threading.Event()
        self.messages = queue.Queue()
        self._controls = []
        self._guard = False
        self._closing = False
        self._u0, self._u1 = 0.0, 1.0
        self._drag = None
        self._plot = None
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._poll_id=self.after(100, self._poll)
        self._refresh_buttons()

    def _control(self, widget, normal="normal"):
        self._controls.append((widget, normal))
        return widget

    def _build(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Treeview", rowheight=25)
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        toolbar = ttk.Frame(root)
        toolbar.pack(fill="x")
        self._control(ttk.Button(toolbar, text="Добавить LAS…", command=self._choose_las)).pack(side="left")
        self._control(ttk.Button(toolbar, text="Выбрать markers.xlsx…", command=self._choose_markers)).pack(side="left", padx=6)
        self.file_label = ttk.Label(toolbar, text="LAS: 0 · Маркеры не загружены")
        self.file_label.pack(side="left", padx=8)
        self._control(ttk.Button(toolbar, text="Сведения о файлах", command=self._file_info)).pack(side="right")

        selector = ttk.LabelFrame(root, text="Входные данные", padding=7)
        selector.pack(fill="x", pady=(8, 5))
        selector.columnconfigure(1, weight=1)
        selector.columnconfigure(3, weight=1)
        self.ref_var, self.target_var = tk.StringVar(), tk.StringVar()
        self.rgroup_var, self.tgroup_var = tk.StringVar(), tk.StringVar()
        ttk.Label(selector, text="Опорная скважина").grid(row=0, column=0, sticky="w", padx=3)
        self.ref_box = self._control(ttk.Combobox(selector, textvariable=self.ref_var, state="readonly", width=30), "readonly")
        self.ref_box.grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Label(selector, text="Целевая скважина").grid(row=0, column=2, sticky="w", padx=3)
        self.target_box = self._control(ttk.Combobox(selector, textvariable=self.target_var, state="readonly", width=30), "readonly")
        self.target_box.grid(row=0, column=3, sticky="ew", padx=5)
        ttk.Label(selector, text="Маркеры опорной").grid(row=1, column=0, sticky="w", padx=3, pady=5)
        self.rgroup_box = self._control(ttk.Combobox(selector, textvariable=self.rgroup_var, state="readonly"), "readonly")
        self.rgroup_box.grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Label(selector, text="Маркеры целевой").grid(row=1, column=2, sticky="w", padx=3)
        self.tgroup_box = self._control(ttk.Combobox(selector, textvariable=self.tgroup_var, state="readonly"), "readonly")
        self.tgroup_box.grid(row=1, column=3, sticky="ew", padx=5)
        for box in (self.ref_box, self.target_box):
            box.bind("<<ComboboxSelected>>", self._pair_changed)
        self.rgroup_box.bind("<<ComboboxSelected>>", self._group_changed)
        self.tgroup_box.bind("<<ComboboxSelected>>", self._group_changed)
        self.metadata = ttk.Label(selector, text="Выберите два LAS-файла. Все глубины отображаются в метрах.", wraplength=1150)
        self.metadata.grid(row=2, column=0, columnspan=4, sticky="w", padx=3)

        settings = ttk.Frame(root)
        settings.pack(fill="x", pady=3)
        curves_frame = ttk.LabelFrame(settings, text="Общие кривые · Ctrl/Shift: выбор", padding=5)
        curves_frame.pack(side="left", fill="y", padx=(0, 7))
        self.curve_list = self._control(tk.Listbox(curves_frame, selectmode="extended", exportselection=False, height=5, width=21))
        self.curve_list.pack(side="left", fill="y")
        scroll = ttk.Scrollbar(curves_frame, command=self.curve_list.yview)
        scroll.pack(side="right", fill="y")
        self.curve_list.configure(yscrollcommand=scroll.set)
        self.curve_list.bind("<<ListboxSelect>>", self._input_changed)
        parameters = ttk.LabelFrame(self.settings_dialog, text="Параметры IDTW", padding=10)
        parameters.pack(fill="x", padx=10, pady=5)
        self.vars = {}
        fields = [("step", "Шаг, м", "0.5"), ("band", "Полоса, м", "50"),
                  ("max_points", "Макс. отсчётов", "2000"), ("gap", "Интерп. разрыв, м", "3"),
                  ("k", "Полуокно K", "0"), ("penalty", "Штраф шага", "0.05")]
        for index, (key, label, default) in enumerate(fields):
            var = self.vars[key] = tk.StringVar(value=default)
            row, col = divmod(index, 3)
            ttk.Label(parameters, text=label).grid(row=row, column=col * 2, sticky="w", padx=(4, 2), pady=2)
            entry = self._control(ttk.Entry(parameters, textvariable=var, width=10))
            entry.grid(row=row, column=col * 2 + 1, sticky="ew", padx=(0, 9), pady=2)
        self.vars["normalize"] = tk.StringVar(value="robust")
        ttk.Label(parameters, text="Нормировка").grid(row=2, column=0, sticky="w", padx=4)
        self._control(ttk.Combobox(parameters, textvariable=self.vars["normalize"], values=("robust", "none"), state="readonly", width=8), "readonly").grid(row=2, column=1, sticky="ew", padx=(0, 9))
        ttk.Label(parameters, text="Полоса — отклонение от линии концов интервалов.").grid(row=2, column=2, columnspan=4, sticky="w", padx=4)
        intervals = ttk.Frame(parameters)
        intervals.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        for i, (key, title) in enumerate((("ref_start", "Опорная MD от"), ("ref_end", "до"), ("target_start", "Целевая MD от"), ("target_end", "до"))):
            ttk.Label(intervals, text=title).grid(row=0, column=2*i, padx=(4, 2))
            self.vars[key] = tk.StringVar()
            self._control(ttk.Entry(intervals, textvariable=self.vars[key], width=9)).grid(row=0, column=2*i+1)
        ttk.Label(intervals, text="м; пусто = весь диапазон").grid(row=0, column=8, padx=8)
        for var in self.vars.values():
            var.trace_add("write", self._input_changed)

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(5, 8))
        self.run_button = ttk.Button(actions, text="Рассчитать корреляцию", command=self._run)
        self.run_button.pack(side="left")
        self.cancel_button = ttk.Button(actions, text="Отмена", command=self._cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(actions, length=180, mode="determinate")
        self.progress.pack(side="left", padx=4)
        self.status = tk.StringVar(value="Ожидание файлов")
        ttk.Label(actions, textvariable=self.status).pack(side="left", padx=8)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)
        compare = ttk.Frame(self.notebook, padding=5)
        report = ttk.Frame(self.notebook, padding=5)
        help_frame = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(compare, text="Сравнение скважин")
        self.notebook.add(report, text="Статистика и маркеры")
        self.notebook.add(help_frame, text="Справка")
        plot_toolbar = ttk.Frame(compare)
        plot_toolbar.pack(fill="x")
        ttk.Label(plot_toolbar, text="Показать кривую:").pack(side="left")
        self.display_var = tk.StringVar()
        self.display_box = ttk.Combobox(plot_toolbar, textvariable=self.display_var, state="readonly", width=18)
        self.display_box.pack(side="left", padx=5)
        self.display_box.bind("<<ComboboxSelected>>", lambda event: self._draw())
        ttk.Button(plot_toolbar, text="Сброс масштаба", command=self._reset_view).pack(side="left", padx=5)
        ttk.Label(plot_toolbar, text="Колесо: глубина · Shift + колесо: масштаб · Наведение: соответствие").pack(side="right", padx=4)
        self.canvas = tk.Canvas(compare, background="#fafbfd", highlightthickness=1, highlightbackground="#d9e0e8")
        self.canvas.pack(fill="both", expand=True, pady=5)
        self.hover = tk.StringVar(value="Цвет закреплён за маркером; пунктир — прогноз, сплошная отметка — известная граница.")
        ttk.Label(compare, textvariable=self.hover).pack(fill="x")
        self.canvas.bind("<Configure>", lambda event: self._draw())
        self.canvas.bind("<MouseWheel>", self._zoom)
        self.canvas.bind("<Button-4>", self._zoom)
        self.canvas.bind("<Button-5>", self._zoom)
        self.canvas.bind("<ButtonPress-1>", self._pan_start)
        self.canvas.bind("<B1-Motion>", self._pan_move)
        self.canvas.bind("<ButtonRelease-1>", lambda event: setattr(self, "_drag", None))
        self.canvas.bind("<Motion>", self._hover)
        self.canvas.bind("<Leave>", lambda event: self.canvas.delete("hover"))
        stats_frame = ttk.Frame(report)
        stats_frame.pack(fill="x", pady=(0, 6))
        self.stats_text = tk.Text(stats_frame, height=7, wrap="word", relief="flat", font=("Segoe UI", 10))
        stats_scroll = ttk.Scrollbar(stats_frame, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=stats_scroll.set)
        stats_scroll.pack(side="right", fill="y")
        self.stats_text.pack(fill="x", expand=True)
        self.stats_text.configure(state="disabled")
        cols = ("marker", "ref_md", "predicted_md", "actual_md", "error", "semblance", "status")
        self.table = ttk.Treeview(report, columns=cols, show="headings", selectmode="browse")
        titles = ("Маркер", "Опорная MD, м", "Прогноз MD, м", "Известная MD, м", "Ошибка, м", "Semblance", "Статус")
        for key, title in zip(cols, titles):
            self.table.heading(key, text=title)
            self.table.column(key, width=110 if key != "status" else 260, minwidth=65, anchor="w" if key in ("marker", "status") else "e")
        table_scroll = ttk.Scrollbar(report, command=self.table.yview)
        self.table.configure(yscrollcommand=table_scroll.set)
        table_scroll.pack(side="right", fill="y")
        self.table.pack(fill="both", expand=True)
        self.table.bind("<<TreeviewSelect>>", self._select_marker)
        self.export_button = ttk.Button(report, text="Экспорт таблицы CSV…", command=self._export, state="disabled")
        self.export_button.pack(anchor="w", pady=6)
        help_text = tk.Text(help_frame, wrap="word", relief="flat", font=("Segoe UI", 11), padx=12, pady=12)
        help_scroll = ttk.Scrollbar(help_frame, command=help_text.yview)
        help_text.configure(yscrollcommand=help_scroll.set)
        help_scroll.pack(side="right", fill="y")
        help_text.pack(fill="both", expand=True)
        help_text.insert("1.0", """ПОРЯДОК РАБОТЫ
1. Добавьте два или несколько LAS-файлов и выберите markers.xlsx.
2. Выберите опорную и целевую скважины, затем общие кривые (Ctrl/Shift для нескольких).
3. Проверьте привязку групп маркеров. Пустая привязка означает отсутствие маркеров. Идентификация по UWI/имени может требовать ручного исправления.
4. При необходимости задайте сопоставимые интервалы MD. Пустая граница берётся из диапазона LAS. Нажмите «Рассчитать корреляцию».
5. Кнопка «Интервалы по маркерам» заполняет MD по диапазону опорных маркеров: запас 10 м для опорной и 100 м для целевой. Запасы и смещение изменяются в отдельном окне «Настройки».
6. Вкладка «Последовательный набор»: добавьте скважины в список, задайте порядок перетаскиванием и проверьте группы маркеров. Для каждой пары можно автоматически определять интервалы. Неопределённые переносы помечаются при распространении дальше.
7. «Открыть результат отдельно»: общий профиль набора, выбор пары и генерации, диапазоны чувствительности, таблица и история ручных правок. Исправление применяется к маркеру; сам путь каротажа не превращается в жёстко привязанный к нему путь.
8. Сохраняйте проект через меню «Проект»: файл .idtw содержит снимок данных, настройки, порядок, генерации и историю. Для восстановления исходные LAS/Excel не требуются.

АЛГОРИТМ И ОГРАНИЧЕНИЯ
Используется локальная ошибка 1 − semblance с преобразованием Гильберта и поиск пути динамическим программированием. Несколько выбранных кривых имеют равные веса. K — полуокно semblance в отсчётах сетки. Semblance является мерой сходства, а не вероятностью правильной корреляции.
DTW фиксирует концы выбранных интервалов: это не автоматический поиск общей стратиграфической части. Полоса в метрах ограничивает отклонение от прямой между концами интервалов. Штраф шага сдерживает растяжение и сжатие.
Нормировка robust меняет амплитуды исходного сигнала по сравнению с формулой Fang; none сохраняет исходные амплитуды. Нормировка не устраняет геологическую неоднозначность.
Длинные пропуски не заполняются интерполяцией; они получают штраф, а перенос маркера в такой участок отклоняется. Максимальный разрыв задаётся в метрах. Не увеличивайте его для сокрытия плохого покрытия.
Число отсчётов ограничивает расход памяти и время; фактический шаг может стать крупнее заданного. Расчёт выполняется в фоновом потоке и может быть отменён.

МАРКЕРЫ И ПРОВЕРКА
Обязательные заголовки markers.xlsx: Маркер, Скважина, UWI, MD, X, Y, Z. MD задаётся в метрах. X/Y/Z сохраняются как метаданные; вертикальная ось сравнения — MD, а не абсолютная отметка Z.
Маркеры опорной скважины переносятся по пути DTW. В обычном и многовариантном режиме известные маркеры целевой используются только для проверки. Ошибка — прогноз минус известная MD. MAE считается только по проверяемым автоматическим переносам; ручные правки исключены. Покрытие нужно оценивать отдельно.
В режиме «Подобрать настройки по разметке» проверяется набор комбинаций параметров. Выбирается минимальная средняя абсолютная ошибка со штрафом за каждый пропуск. Используются исходные и подтверждённые пользователем контрольные маркеры. Это подбор на известных ответах; его ошибка не является независимой проверкой и лучший вариант не является доказанным глобальным оптимумом. Лучшие параметры автоматически записываются в окно настроек после завершения подбора и сохраняются при смене скважин. Кнопка в окне результата позволяет повторно применить параметры выбранного результата. Сохранение проекта сохраняет и текущие настройки.
Многовариантный режим выбирает реальный путь из наиболее поддержанной группы решений, а не усредняет разные границы. P10–P90 показывает чувствительность внутри этой группы; альтернативные группы отображаются отдельно. Поддержка не является вероятностью правильности. Провальные генерации учитываются в знаменателе поддержки.
Ручную MD можно изменить и подтвердить двойным щелчком в таблице отдельного окна. История сохраняется. Последующие пары после такой правки устаревают; кнопка «Пересчитать после этой скважины» использует исправленные границы как исходные для переноса дальше. Подтверждённые правки сохраняются при пересчётах; конфликты отмечаются явно.
Для неизвестного UWI выбирайте группу вручную. При отсутствии маркеров по-прежнему можно сравнить каротаж и путь корреляции.

ГРАФИК
Дорожки показывают исходные значения выбранной кривой с отдельными шкалами амплитуды. Одинаковое имя маркера имеет одинаковый цвет. В основном окне пунктир — прогноз, сплошная отметка — контроль; серые линии показывают примеры соответствий пути. В отдельном окне точка означает контроль, ромб — ручное подтверждение; сомнительные линии пунктирные, цветные зоны — P10–P90.
Наведение показывает соответствие глубин и показатели. Колесо перемещает глубину, Shift + колесо масштабирует около указателя, перетаскивание перемещает окно. Основной график использует относительные окна двух скважин, отдельное окно — общую шкалу MD для всего профиля. Масштаб не меняет результаты расчёта.

Экспорт создаёт отдельный CSV с UTF-8 BOM для открытия в Excel. Исходные LAS и XLSX не изменяются.
""")
        help_text.configure(state="disabled")

    @staticmethod
    def _fmt(value, digits=3):
        if value is None:
            return "—"
        if isinstance(value, (float, np.floating)):
            return f"{value:.{digits}f}" if math.isfinite(value) else "—"
        return str(value)

    def _current(self):
        return self.wells.get(self.ref_var.get()), self.wells.get(self.target_var.get())

    def set_dataset(self, wells, markers):
        """Populate data from the main thread (also used by --demo/--gui-smoke)."""
        if self.busy:
            raise RuntimeError("Нельзя менять данные во время расчёта.")
        self.wells.clear()
        for well in wells:
            label = f"{well.name or Path(well.path).stem} | {well.uwi or 'без UWI'} | {Path(well.path).name}"
            if label in self.wells:
                label += f" ({len(self.wells) + 1})"
            self.wells[label] = well
        self.markers = list(markers)
        self.groups = marker_groups(self.markers)
        self.assignments.clear()
        self.marker_path = "данные в памяти"
        labels = list(self.wells)
        self.ref_box.configure(values=labels)
        self.target_box.configure(values=labels)
        self.ref_var.set(labels[0] if labels else "")
        self.target_var.set(labels[1] if len(labels) > 1 else "")
        self._pair_changed()
        self._update_files()

    def _selected_curves(self):
        return [self.curve_list.get(i) for i in self.curve_list.curselection()]

    def _refresh_buttons(self):
        ref, target = self._current()
        valid = ref is not None and target is not None and ref.path != target.path and bool(self._selected_curves())
        self.run_button.configure(state="normal" if valid and not self.busy else "disabled")
        self.cancel_button.configure(state="normal" if self.busy else "disabled")
        self.export_button.configure(state="normal" if self.result is not None and not self.busy else "disabled")

    def _set_busy(self, busy):
        self.busy = busy
        for widget, normal in self._controls:
            widget.configure(state="disabled" if busy else normal)
        self._refresh_buttons()

    def _input_changed(self, *args):
        if self._guard or self.busy:
            return
        self.result = None
        self._plot = None
        self.hover.set("После расчёта наведите указатель на дорожку для просмотра соответствия глубин.")
        for item in self.table.get_children():
            self.table.delete(item)
        self._set_stats("")
        self.display_box.configure(values=())
        self.display_var.set("")
        self.status.set("Параметры изменены. Выполните расчёт.")
        self._refresh_buttons()
        self._draw()

    def _pair_changed(self, event=None):
        self._guard = True
        ref, target = self._current()
        old = self._selected_curves()
        common = sorted(set(ref.curves).intersection(target.curves)) if ref and target else []
        self.curve_list.delete(0, "end")
        for curve in common:
            self.curve_list.insert("end", curve)
        chosen = [i for i, curve in enumerate(common) if curve in old]
        for i in chosen or ([0] if common else []):
            self.curve_list.selection_set(i)
        choices = [""] + sorted(self.groups)
        for box, var, well in ((self.rgroup_box, self.rgroup_var, ref), (self.tgroup_box, self.tgroup_var, target)):
            box.configure(values=choices)
            group = self.assignments.get(well.path, match_group(well, self.groups)) if well else ""
            var.set(group if group in self.groups else "")
        if ref and target:
            self.metadata.configure(text=f"Опорная: {ref.name} · UWI {ref.uwi or '—'} · {ref.depth[0]:.2f}–{ref.depth[-1]:.2f} м | Целевая: {target.name} · UWI {target.uwi or '—'} · {target.depth[0]:.2f}–{target.depth[-1]:.2f} м")
        self._guard = False
        self._input_changed()

    def _group_changed(self, event=None):
        for well, var in zip(self._current(), (self.rgroup_var, self.tgroup_var)):
            if well:
                self.assignments[well.path] = var.get()
        self._input_changed()

    def _start_worker(self, work):
        self.cancel_event.clear()
        self.progress["value"] = 0
        self._set_busy(True)
        def run():
            try:
                kind, payload = work()
                self.messages.put((kind, payload))
            except Exception as exc:
                self.messages.put(("error", str(exc)))
            finally:
                self.messages.put(("done", None))
        threading.Thread(target=run, daemon=True, name="idtw-worker").start()

    def _choose_las(self):
        paths = filedialog.askopenfilenames(parent=self, title="Выберите LAS-файлы", filetypes=[("LAS", "*.las *.LAS"), ("Все файлы", "*.*")])
        if not paths:
            return
        self.status.set("Чтение LAS…")
        def work():
            loaded, errors = [], []
            for i, path in enumerate(paths):
                if self.cancel_event.is_set():
                    break
                try:
                    loaded.append(load_las(path))
                except Exception as exc:
                    errors.append(f"{path}\n{exc}")
                self.messages.put(("progress", ((i + 1) / len(paths), f"LAS: {i + 1}/{len(paths)}")))
            return "las", (loaded, errors)
        self._start_worker(work)

    def _choose_markers(self):
        path = filedialog.askopenfilename(parent=self, title="Выберите markers.xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.status.set("Чтение маркеров…")
            self._start_worker(lambda: ("markers", (path, load_markers(path))))

    def _poll(self):
        if self._closing:
            return
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "progress":
                    self.progress["value"] = max(0, min(100, payload[0] * 100))
                    self.status.set(payload[1])
                elif kind == "done":
                    self._set_busy(False)
                elif kind == "error":
                    self.status.set("Расчёт отменён" if self.cancel_event.is_set() else "Ошибка")
                    if not self.cancel_event.is_set():
                        messagebox.showerror("IDTW", payload, parent=self)
                elif kind == "las":
                    self._set_busy(False)
                    loaded, errors = payload
                    for well in loaded:
                        for label, old in list(self.wells.items()):
                            if old.path == well.path:
                                del self.wells[label]
                        label = f"{well.name or Path(well.path).stem} | {well.uwi or 'без UWI'} | {Path(well.path).name}"
                        if label in self.wells:
                            label += f" ({len(self.wells) + 1})"
                        self.wells[label] = well
                    labels = list(self.wells)
                    self.ref_box.configure(values=labels)
                    self.target_box.configure(values=labels)
                    if self.ref_var.get() not in self.wells and labels:
                        self.ref_var.set(labels[0])
                    if self.target_var.get() not in self.wells and len(labels) > 1:
                        self.target_var.set(labels[1])
                    self._pair_changed()
                    self._update_files()
                    self.status.set(f"Загружено LAS: {len(loaded)}")
                    if errors:
                        messagebox.showwarning("Некоторые LAS не прочитаны", "\n\n".join(errors), parent=self)
                elif kind == "markers":
                    self._set_busy(False)
                    if self.cancel_event.is_set():
                        self.status.set("Загрузка маркеров отменена")
                        continue
                    self.marker_path, self.markers = payload
                    self.groups = marker_groups(self.markers)
                    self.assignments.clear()
                    self._pair_changed()
                    self._update_files()
                    self.status.set(f"Загружено маркеров: {len(self.markers)}")
                elif kind == "result":
                    if self.cancel_event.is_set():
                        self.status.set("Расчёт отменён")
                    else:
                        self.result = payload
                        self._show_result()
                elif kind in ("analysis", "project", "saved"):
                    if not self.cancel_event.is_set():
                        self._handle_extra(kind, payload)
        except queue.Empty:
            pass
        self._poll_id=self.after(80, self._poll)

    def _update_files(self):
        marker_text = f"маркеров: {len(self.markers)} ({Path(self.marker_path).name})" if hasattr(self, "marker_path") else "маркеры не загружены"
        self.file_label.configure(text=f"LAS: {len(self.wells)} · {marker_text}")

    def _read_params(self):
        values = {}
        for key, var in self.vars.items():
            text = var.get().strip().replace(",", ".")
            if key == "normalize":
                values[key] = text
            elif key in ("ref_start", "ref_end", "target_start", "target_end"):
                values[key] = float(text) if text else None
            elif key in ("max_points", "k"):
                values[key] = int(text)
            else:
                values[key] = float(text)
        if any(isinstance(v, float) and not math.isfinite(v) for v in values.values()):
            raise ValueError("Параметры должны быть конечными числами.")
        if values["step"] <= 0 or values["band"] <= 0 or values["gap"] <= 0 or values["penalty"] < 0:
            raise ValueError("Шаг, полоса и разрыв должны быть положительными; штраф — неотрицательным.")
        if not 50 <= values["max_points"] <= 4000 or not 0 <= values["k"] <= 10:
            raise ValueError("Макс. отсчётов: 50–4000; K: 0–10.")
        for prefix in ("ref", "target"):
            start, end = values[prefix + "_start"], values[prefix + "_end"]
            if start is not None and end is not None and start >= end:
                raise ValueError("Начало интервала MD должно быть меньше конца.")
        return Params(**values)

    def _run(self):
        ref, target = self._current()
        curves = self._selected_curves()
        if not ref or not target or ref.path == target.path or not curves:
            messagebox.showerror("Входные данные", "Выберите две разные скважины и хотя бы одну общую кривую.", parent=self)
            return
        try:
            params = self._read_params()
        except (ValueError, TypeError) as exc:
            messagebox.showerror("Параметры", str(exc), parent=self)
            return
        rmarkers = list(self.groups.get(self.rgroup_var.get(), []))
        tmarkers = list(self.groups.get(self.tgroup_var.get(), []))
        self._input_changed()
        self.status.set("Подготовка кривых…")
        def work():
            result = correlate(ref, target, curves, params, rmarkers, tmarkers, self.cancel_event,
                               lambda fraction, message: self.messages.put(("progress", (fraction, message))))
            return "result", result
        self._start_worker(work)

    def _cancel(self):
        self.cancel_event.set()
        self.status.set("Отмена…")

    def _show_result(self):
        result = self.result
        self.display_box.configure(values=result.curves)
        self.display_var.set(result.curves[0])
        self._u0, self._u1 = 0.0, 1.0
        details = [f"{key}: {self._fmt(value)}" for key, value in result.stats.items()]
        if not result.stats.get("Предупреждения"):
            details += list(result.notes)
        details.append("Semblance — мера сходства, не вероятность. " +
                       ("Целевые маркеры использованы для ПОДБОРА; ошибка не независимая."
                        if result.provenance.get('mode')=='tuned' or result.provenance.get('calibrated_on_target') else "Целевые маркеры использованы только для проверки."))
        self._set_stats("\n".join(details))
        for item in self.table.get_children():
            self.table.delete(item)
        for i, row in enumerate(result.rows):
            self.table.insert("", "end", iid=str(i), values=[self._fmt(row.get(key)) for key in self.table["columns"]])
        self.progress["value"] = 100
        self.status.set("Корреляция рассчитана")
        self.notebook.select(0)
        self._draw()

    def _set_stats(self, text):
        self.stats_text.configure(state="normal")
        self.stats_text.delete("1.0", "end")
        self.stats_text.insert("1.0", text)
        self.stats_text.configure(state="disabled")

    def _file_info(self):
        dialog = tk.Toplevel(self)
        dialog.title("Файлы и метаданные")
        dialog.geometry("900x550")
        text = tk.Text(dialog, wrap="word", padx=10, pady=10)
        scroll = ttk.Scrollbar(dialog, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        for well in self.wells.values():
            text.insert("end", f"{well.name} · UWI: {well.uwi or '—'}\n{well.path}\nMD: {well.depth[0]:.3f}–{well.depth[-1]:.3f} м; отсчётов: {len(well.depth)}\nКривые: {', '.join(well.curves)}\n")
            for note in well.notes:
                text.insert("end", f"  • {note}\n")
            text.insert("end", "\n")
        text.insert("end", f"Маркеры: {getattr(self, 'marker_path', 'не загружены')}\nСтрок: {len(self.markers)}; групп: {len(self.groups)}")
        text.configure(state="disabled")

    def _export(self):
        if self.result is None:
            return
        path = filedialog.asksaveasfilename(parent=self, title="Экспорт маркеров", defaultextension=".csv", initialfile="idtw_markers.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        keys = ["marker", "ref_md", "predicted_md", "actual_md", "error", "semblance", "status"]
        def safe_cell(value):
            # Keep text cells as text when the CSV is opened in Excel.
            if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')):
                return "'" + value
            return value
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as stream:
                writer = csv.writer(stream, delimiter=";")
                writer.writerow(["Опорная скважина", "Целевая скважина"] + keys)
                for row in self.result.rows:
                    writer.writerow([safe_cell(v) for v in [self.result.ref.name, self.result.target.name]
                                     + [row.get(key, "") for key in keys]])
            self.status.set(f"Сохранено: {path}")
        except OSError as exc:
            messagebox.showerror("Экспорт", str(exc), parent=self)

    def _reset_view(self):
        self._u0, self._u1 = 0.0, 1.0
        self._draw()

    def _bounded_view(self, start, span):
        span = max(0.005, min(1.0, span))
        self._u0 = max(0.0, min(1.0 - span, start))
        self._u1 = self._u0 + span
        self._draw()

    def _zoom(self, event):
        if not self._plot:
            return
        top, height = self._plot["top"], self._plot["height"]
        relative = min(1.0, max(0.0, (event.y - top) / height))
        span = self._u1 - self._u0
        anchor = self._u0 + relative * span
        direction = 1 if getattr(event, "num", None) == 4 or getattr(event, "delta", 0) > 0 else -1
        if not getattr(event, 'state', 0) & 0x0001:
            self._bounded_view(self._u0 - direction * span * .10, span)
            return
        new_span = min(1.0, max(0.005, span * (0.8 if direction > 0 else 1.25)))
        self._bounded_view(anchor - relative * new_span, new_span)

    def _pan_start(self, event):
        self._drag = (event.y, self._u0, self._u1 - self._u0)

    def _pan_move(self, event):
        if self._drag and self._plot:
            y, start, span = self._drag
            self._bounded_view(start - (event.y - y) * span / self._plot["height"], span)

    def _y(self, depth, target=False):
        z = self.result.zt if target else self.result.zr
        unit = (depth - z[0]) / max(1e-12, z[-1] - z[0])
        return self._plot["top"] + (unit - self._u0) / (self._u1 - self._u0) * self._plot["height"]

    def _draw(self):
        c = self.canvas
        c.delete("all")
        width, height = c.winfo_width(), c.winfo_height()
        if width < 50 or height < 50:
            return
        if self.result is None:
            self._plot = None
            c.create_text(width / 2, height / 2, text="Выберите LAS-файлы, кривые и выполните расчёт.\nЗдесь появятся каротаж, линии корреляции и маркеры.", font=("Segoe UI", 13), fill="#576574", justify="center")
            return
        top, bottom = 48, height - 34
        left = (67, width * 0.37)
        right = (width * 0.65, width - 60)
        self._plot = dict(top=top, height=max(20, bottom - top), left=left, right=right)
        curve = self.display_var.get()
        result = self.result
        for target, well, bounds, z in ((False, result.ref, left, result.zr), (True, result.target, right, result.zt)):
            lo, hi = bounds
            c.create_rectangle(lo, top, hi, bottom, fill="#ffffff", outline="#a6b6c7")
            c.create_text((lo + hi) / 2, 15, text=well.name[:38], fill="#24354b", font=("Segoe UI", 11, "bold"))
            c.create_text((lo + hi) / 2, 33, text=f"{curve} ({well.units.get(curve, '')}) · MD, м", fill="#536779")
            for fraction in np.linspace(0, 1, 7):
                u = self._u0 + fraction * (self._u1 - self._u0)
                depth = z[0] + u * (z[-1] - z[0])
                y = top + fraction * (bottom - top)
                c.create_line(lo, y, hi, y, fill="#e6ebf2")
                c.create_text(lo - 5, y, text=f"{depth:.1f}", anchor="e", fill="#536779", font=("Segoe UI", 9))
            values = well.curves.get(curve)
            if values is None:
                continue
            depth = well.depth
            valid_all = np.isfinite(values) & (depth >= z[0]) & (depth <= z[-1])
            if not valid_all.any():
                continue
            xmin, xmax = float(np.min(values[valid_all])), float(np.max(values[valid_all]))
            if xmax <= xmin:
                xmin, xmax = xmin - 0.5, xmax + 0.5
            pad = 0.03 * (xmax - xmin)
            xmin, xmax = xmin - pad, xmax + pad
            d0 = z[0] + self._u0 * (z[-1] - z[0])
            d1 = z[0] + self._u1 * (z[-1] - z[0])
            a = max(0, int(np.searchsorted(depth, d0)) - 1)
            b = min(len(depth), int(np.searchsorted(depth, d1, side="right")) + 1)
            indices = np.arange(a, b)
            finite = np.isfinite(values[indices])
            # Split before decimation, so a null interval never becomes a drawn bridge.
            keep = indices[finite]
            breaks = np.diff(keep) > 1
            if result.params.gap > 0:
                breaks |= np.diff(depth[keep]) > result.params.gap
            blocks = np.split(keep, np.where(breaks)[0] + 1)
            for block in blocks:
                if len(block) < 2:
                    continue
                stride = max(1, len(block) // 1800)
                block = np.unique(np.append(block[::stride], block[-1]))
                coords = []
                for index in block:
                    px = lo + (values[index] - xmin) / (xmax - xmin) * (hi - lo)
                    px = max(lo, min(hi, px))
                    py = max(top, min(bottom, self._y(float(depth[index]), target)))
                    coords.extend((float(px), float(py)))
                c.create_line(*coords, fill="#2166ac", width=1.2)
            c.create_text(lo, bottom + 15, text=f"{xmin:.3g}", anchor="w", fill="#536779")
            c.create_text(hi, bottom + 15, text=f"{xmax:.3g}", anchor="e", fill="#536779")
        for i in np.linspace(0, len(result.zr) - 1, min(28, len(result.zr)), dtype=int):
            if not np.isfinite(result.mapped[i]) or not np.isfinite(result.similarity[i]):
                continue
            yr, yt = self._y(result.zr[i]), self._y(result.mapped[i], True)
            if top <= yr <= bottom and top <= yt <= bottom:
                c.create_line(left[1], yr, right[0], yt, fill="#d6dce3")
        for row in result.rows:
            if row.get("ref_md") is None:
                continue
            yr = self._y(row["ref_md"])
            color=marker_color(row['marker'])
            pred, actual = row.get("predicted_md"), row.get("actual_md")
            if top <= yr <= bottom:
                c.create_line(left[0], yr, left[1], yr, fill=color, dash=(4, 3))
                c.create_text(left[1] + 3, yr - 2, text=str(row["marker"])[:24], anchor="sw", fill=color, font=("Segoe UI", 9))
            if pred is not None and np.isfinite(pred):
                yt = self._y(pred, True)
                if top <= yt <= bottom:
                    c.create_line(right[0], yt, right[1], yt, fill=color, dash=(4, 3))
                if top <= yr <= bottom and top <= yt <= bottom:
                    c.create_line(left[1], yr, right[0], yt, fill=color, width=1.8)
            if actual is not None and np.isfinite(actual):
                ya = self._y(actual, True)
                if top <= ya <= bottom:
                    c.create_line(right[0], ya, right[1], ya, fill=color, width=2)
                    c.create_text(right[1] - 3, ya - 2, text=str(row["marker"])[:24], anchor="se", fill=color, font=("Segoe UI", 9))
        shown = {_key(row['marker']) for row in result.rows if row.get('actual_md') is not None}
        for marker in self.groups.get(self.tgroup_var.get(), []):
            if _key(marker.name) in shown:
                continue
            ya = self._y(marker.md, True)
            if top <= ya <= bottom:
                c.create_line(right[0], ya, right[1], ya, fill=marker_color(marker.name), width=2)
                c.create_text(right[1]-3, ya-2, text=marker.name[:24], anchor="se", fill=marker_color(marker.name), font=("Segoe UI", 9))

    def _hover(self, event):
        if not self._plot or self._drag:
            return
        c, p, result = self.canvas, self._plot, self.result
        c.delete("hover")
        if not p["top"] <= event.y <= p["top"] + p["height"]:
            return
        u = self._u0 + (event.y - p["top"]) / p["height"] * (self._u1 - self._u0)
        if event.x >= p["right"][0]:
            depth = result.zt[0] + u * (result.zt[-1] - result.zt[0])
            indices = np.flatnonzero(np.isfinite(result.mapped))
            if not len(indices):
                return
            index = int(indices[np.argmin(np.abs(result.mapped[indices] - depth))])
        else:
            depth = result.zr[0] + u * (result.zr[-1] - result.zr[0])
            index = min(len(result.zr) - 1, int(np.searchsorted(result.zr, depth)))
        source, mapped = result.zr[index], result.mapped[index]
        score = result.similarity[index]
        self.hover.set(f"Опорная MD: {source:.3f} м · Целевая MD: {self._fmt(mapped)} м · Semblance по выбранным кривым: {self._fmt(score)} (не вероятность)")
        yr = self._y(source)
        source_visible = p["top"] <= yr <= p["top"] + p["height"]
        if source_visible:
            c.create_line(p["left"][0], yr, p["left"][1], yr, fill="#9b299b", width=2, tags="hover")
        if np.isfinite(mapped):
            yt = self._y(mapped, True)
            if p["top"] <= yt <= p["top"] + p["height"]:
                if source_visible:
                    c.create_line(p["left"][1], yr, p["right"][0], yt, fill="#9b299b", width=2, tags="hover")
                c.create_line(p["right"][0], yt, p["right"][1], yt, fill="#9b299b", width=2, tags="hover")

    def _select_marker(self, event=None):
        if self.result is None or not self.table.selection():
            return
        row = self.result.rows[int(self.table.selection()[0])]
        if row.get("ref_md") is None:
            return
        z = self.result.zr
        unit = (row["ref_md"] - z[0]) / max(1e-12, z[-1] - z[0])
        span = min(0.25, self._u1 - self._u0)
        self._bounded_view(unit - span / 2, span)

    def _close(self):
        self.cancel_event.set()
        self._closing = True
        if getattr(self,'_poll_id',None):
            self.after_cancel(self._poll_id)
        self.destroy()


class IDTWApp(PairApp):
    """Application controller for pair, sequence, ensemble and project workflows."""
    def __init__(self):
        self.bundles = []
        self.sequence_paths = []
        self.result_windows = []
        self.project_path = None
        self.manual_picks = {}
        self.calibration_context = {}
        self.dirty = False
        super().__init__()
        self.title('IDTW — корреляция скважин и контроль вариантов')
        self.geometry('1280x900')

    def _build(self):
        self.settings_dialog = tk.Toplevel(self)
        self.settings_dialog.title('Настройки IDTW')
        self.settings_dialog.geometry('1010x660')
        self.settings_dialog.protocol('WM_DELETE_WINDOW',self.settings_dialog.withdraw)
        self.settings_dialog.withdraw()
        super()._build()
        settings = ttk.LabelFrame(self.settings_dialog,text='Автоинтервалы и генерации',padding=10)
        settings.pack(fill='x',padx=10,pady=5)
        self.config_vars = {}
        options = [('ref_margin','Запас опорной, м','10'),('target_margin','Запас целевой, м','100'),
                   ('shift','Смещение целевой, м','0'),('count','Генераций (2–64)','16'),
                   ('seed','Seed генераций','42'),('tolerance','Допуск согласия/ошибки, м','2')]
        for i,(key,label,value) in enumerate(options):
            self.config_vars[key] = tk.StringVar(value=value)
            ttk.Label(settings,text=label).grid(row=i//2,column=(i%2)*2,sticky='w',padx=5,pady=4)
            self._control(ttk.Entry(settings,textvariable=self.config_vars[key],width=16)).grid(row=i//2,column=(i%2)*2+1,padx=8)
        self.config_vars['ensemble'] = tk.BooleanVar(value=False)
        self.config_vars['auto'] = tk.BooleanVar(value=True)
        self._control(ttk.Checkbutton(settings,text='Многовариантный расчёт (для пары и набора)',variable=self.config_vars['ensemble'])).grid(row=3,column=0,columnspan=4,sticky='w')
        self._control(ttk.Checkbutton(settings,text='Автоинтервалы для каждой пары последовательного набора',variable=self.config_vars['auto'])).grid(row=4,column=0,columnspan=4,sticky='w')
        for var in self.config_vars.values():
            var.trace_add('write',self._input_changed)
        explanation = ('Генератор проверяет полосу ×0.5/1/1.5, шаг ×0.75/1/1.25, K=0/1/2/заданный, '
                       'штраф 0/0.025/0.075/0.15/заданный, набор кривых и исключение одного канала. '
                       'Нормировка и интервалы фиксированы. Seed обеспечивает повторяемость.\n\n'
                       'Консенсус: группировка целых путей по RMS расхождения глубин; выбирается реальный '
                       'представитель наиболее поддержанной группы. P10–P90 — чувствительность внутри группы, '
                       'не доверительный интервал. Альтернативы показываются отдельно.\n\n'
                       'Подбор: средняя абсолютная ошибка по общим контрольным маркерам; за отсутствие прогноза '
                       'штраф = длина целевого интервала + допуск. Это лучший из проверенных вариантов, '
                       'не глобальный оптимум. Ошибки подбора не являются независимой оценкой точности.')
        ttk.Label(self.settings_dialog,text=explanation,wraplength=955,justify='left').pack(fill='x',padx=15,pady=10)
        self._control(ttk.Button(self.settings_dialog,text='Восстановить настройки',command=self._defaults)).pack(side='left',padx=15,pady=8)
        ttk.Button(self.settings_dialog,text='Закрыть',command=self.settings_dialog.withdraw).pack(side='right',padx=15,pady=8)
        # Add a compact toolbar; heavy settings remain in a separate window.
        menu = tk.Menu(self)
        project_menu = tk.Menu(menu,tearoff=False)
        project_menu.add_command(label='Открыть проект…',command=self._open_project)
        project_menu.add_command(label='Сохранить проект…',command=self._save_project)
        menu.add_cascade(label='Проект',menu=project_menu)
        menu.add_command(label='Настройки…',command=self._settings)
        self.configure(menu=menu)
        extra = ttk.Frame(self.notebook.master)
        extra.pack(fill='x',before=self.notebook,pady=5)
        self._control(ttk.Button(extra,text='Интервалы по маркерам',command=self._auto_depth)).pack(side='left',padx=3)
        self._control(ttk.Button(extra,text='Настройки…',command=self._settings)).pack(side='left',padx=3)
        self.tune_button = self._control(ttk.Button(extra,text='Подобрать настройки по разметке',command=lambda:self._run_mode('tune')))
        self.tune_button.pack(side='left',padx=3)
        self.window_button = ttk.Button(extra,text='Открыть результат отдельно',command=self._open_result)
        self.window_button.pack(side='left',padx=3)
        self._control(ttk.Button(extra,text='Сохранить проект…',command=self._save_project)).pack(side='right',padx=3)
        seq = ttk.Frame(self.notebook,padding=10)
        self.notebook.insert(1,seq,text='Последовательный набор')
        ttk.Label(seq,text='Добавьте скважины в правый список. Перетаскивайте строки для изменения порядка А → Б → В.').pack(anchor='w')
        lists = ttk.Frame(seq)
        lists.pack(fill='both',expand=True,pady=8)
        left = ttk.Frame(lists)
        left.pack(side='left',fill='both',expand=True)
        right = ttk.Frame(lists)
        right.pack(side='right',fill='both',expand=True)
        ttk.Label(left,text='Загруженные скважины').pack(anchor='w')
        self.available_list = self._control(tk.Listbox(left,selectmode='extended',exportselection=False,height=7,width=45))
        self.available_list.pack(fill='both',expand=True)
        self._control(ttk.Button(left,text='Добавить выбранные →',command=self._sequence_add)).pack(anchor='e',pady=4)
        ttk.Label(right,text='Порядок расчёта — перетащите строку').pack(anchor='w')
        self.sequence_list = self._control(tk.Listbox(right,exportselection=False,height=7,width=52))
        self.sequence_list.pack(fill='both',expand=True,padx=(12,0))
        self.sequence_list.bind('<ButtonPress-1>',self._drag_sequence_start)
        self.sequence_list.bind('<B1-Motion>',self._drag_sequence)
        self.sequence_list.bind('<<ListboxSelect>>',self._sequence_selection)
        self._control(ttk.Button(right,text='Удалить выбранную',command=self._sequence_remove)).pack(anchor='e',pady=4)
        assignment = ttk.Frame(seq)
        assignment.pack(fill='x',pady=4)
        ttk.Label(assignment,text='Группа маркеров выбранной скважины:').pack(side='left')
        self.sequence_group = tk.StringVar()
        self.sequence_group_box = self._control(ttk.Combobox(assignment,textvariable=self.sequence_group,state='readonly',width=60),'readonly')
        self.sequence_group_box.pack(side='left',padx=8)
        self.sequence_group_box.bind('<<ComboboxSelected>>',self._assign_sequence_group)
        self.sequence_button = self._control(ttk.Button(seq,text='Рассчитать последовательный набор',command=lambda:self._run_mode('sequence')))
        self.sequence_button.pack(anchor='w',pady=6)
        ttk.Label(seq,text='Исходные маркеры берутся из первой скважины. Затем используются перенесённые или вручную подтверждённые границы.\n'
                  'Контрольные маркеры остальных скважин служат для оценки. Неопределённость предыдущих переносов отмечается в статусе.',wraplength=1100).pack(anchor='w')
        self.sequence_status = ttk.Label(seq,text='')
        self.sequence_status.pack(anchor='w',pady=5)

    def _settings(self):
        self.settings_dialog.deiconify()
        self.settings_dialog.lift()

    def _defaults(self):
        self._guard=True
        for key,value in asdict(Params()).items():
            self.vars[key].set('' if value is None else str(value))
        for key,value in dict(ref_margin='10',target_margin='100',shift='0',count='16',seed='42',tolerance='2',ensemble=False,auto=True).items():
            self.config_vars[key].set(value)
        self._guard=False
        self._input_changed()

    def _read_config(self):
        result = {}
        for key,var in self.config_vars.items():
            value = var.get()
            result[key] = bool(value) if key in ('ensemble','auto') else (int(value) if key in ('count','seed') else float(str(value).replace(',','.')))
        if not 2<=result['count']<=64 or result['seed']<0:
            raise ValueError('Генераций: 2–64; seed — неотрицательное целое.')
        if not all(np.isfinite(result[k]) for k in ('ref_margin','target_margin','shift','tolerance')):
            raise ValueError('Параметры должны быть конечными.')
        if min(result['ref_margin'],result['target_margin'])<0 or result['tolerance']<=0:
            raise ValueError('Запасы ≥ 0, допуск > 0.')
        return result

    def _auto_depth(self):
        try:
            ref,target = self._current()
            if not ref or not target:
                raise ValueError('Выберите опорную и целевую скважины.')
            config = self._read_config()
            p = auto_intervals(ref,target,self._groups_snapshot().get(ref.path,[]),self._read_params(),
                               config['ref_margin'],config['target_margin'],config['shift'])
            self._guard=True
            for key in ('ref_start','ref_end','target_start','target_end'):
                self.vars[key].set(f'{getattr(p,key):.6g}')
            self._guard=False
            self._input_changed()
            self.status.set(f'Интервалы: {p.ref_start:.1f}–{p.ref_end:.1f} → {p.target_start:.1f}–{p.target_end:.1f} м')
            self._settings()
        except (ValueError,TypeError) as error:
            self._guard=False
            messagebox.showerror('Автоинтервалы',str(error),parent=self)

    def _input_changed(self,*args):
        if self._guard or self.busy:
            return
        for bundle in self.bundles:
            bundle.result.provenance['settings_changed']=True
        self.dirty=True
        super()._input_changed(*args)
        self._refresh_result_windows()

    def _refresh_buttons(self):
        super()._refresh_buttons()
        if hasattr(self,'window_button'):
            self.window_button.configure(state='normal' if self.bundles and not self.busy else 'disabled')

    def _update_files(self):
        super()._update_files()
        if not hasattr(self,'available_list'):
            return
        self.available_list.delete(0,'end')
        for label in self.wells:
            self.available_list.insert('end',label)
        valid = {w.path for w in self.wells.values()}
        self.sequence_paths=[p for p in self.sequence_paths if p in valid]
        self._sequence_render()

    def _well_by_path(self,path):
        return next(w for w in self.wells.values() if w.path==path)

    def _sequence_render(self):
        self.sequence_list.delete(0,'end')
        for i,path in enumerate(self.sequence_paths):
            well=self._well_by_path(path)
            self.sequence_list.insert('end',f'{i+1}. {well.name} | {well.uwi or "без UWI"}')
        self.sequence_group_box.configure(values=['']+sorted(self.groups))

    def _sequence_add(self):
        labels=list(self.wells)
        for index in self.available_list.curselection():
            path=self.wells[labels[index]].path
            if path not in self.sequence_paths:
                self.sequence_paths.append(path)
        self._sequence_render()
        self._input_changed()

    def _sequence_remove(self):
        if self.sequence_list.curselection():
            del self.sequence_paths[self.sequence_list.curselection()[0]]
            self._sequence_render()
            self._input_changed()

    def _drag_sequence_start(self,event):
        self._sequence_drag_index=self.sequence_list.nearest(event.y)

    def _drag_sequence(self,event):
        if self.busy or not self.sequence_paths:
            return
        old=getattr(self,'_sequence_drag_index',0)
        new=self.sequence_list.nearest(event.y)
        if old!=new and 0<=old<len(self.sequence_paths):
            self.sequence_paths.insert(new,self.sequence_paths.pop(old))
            self._sequence_drag_index=new
            self._sequence_render()
            self.sequence_list.selection_set(new)
            self._sequence_selection()
            self._input_changed()

    def _sequence_selection(self,event=None):
        selected=self.sequence_list.curselection()
        if selected:
            well=self._well_by_path(self.sequence_paths[selected[0]])
            self.sequence_group.set(self.assignments.get(well.path,match_group(well,self.groups)))

    def _assign_sequence_group(self,event=None):
        selected=self.sequence_list.curselection()
        if selected:
            self.assignments[self.sequence_paths[selected[0]]]=self.sequence_group.get()
            self._pair_changed()

    def _groups_snapshot(self):
        mapping={w.path:list(self.groups.get(self.assignments.get(w.path,match_group(w,self.groups)),[])) for w in self.wells.values()}
        for well,var in zip(self._current(),(self.rgroup_var,self.tgroup_var)):
            if well:
                mapping[well.path]=list(self.groups.get(var.get(),[]))
        # User-confirmed picks are interpreted labels, not automatic pseudo-labels.
        for well in self.wells.values():
            values={_key(m.name):m for m in mapping[well.path]}
            for name,row in self.manual_picks.get(well.path,{}).items():
                if row.get('confirmed') and row.get('predicted_md') is not None:
                    values[name]=Marker(row['marker'],well.name,well.uwi,row['predicted_md'])
            mapping[well.path]=sorted(values.values(),key=lambda m:m.md)
        return mapping

    def _run(self):
        self._run_mode('pair')

    def _run_mode(self,mode,start_index=None):
        if self.busy:
            return
        try:
            params,config=self._read_params(),self._read_config()
            curves=self._selected_curves()
            if not curves:
                raise ValueError('Выберите кривые для расчёта.')
            mapping=self._groups_snapshot()
            manual_picks=copy.deepcopy(self.manual_picks)
            calibration_context=copy.deepcopy(self.calibration_context)
            prefix=[]
            initial=None
            if mode=='sequence':
                if start_index is None:
                    wells=[self._well_by_path(p) for p in self.sequence_paths]
                else:
                    prefix=self.bundles[:start_index+1]
                    prefix_paths=[prefix[0].result.ref.path]+[b.result.target.path for b in prefix]
                    if self.sequence_paths[:len(prefix_paths)]==prefix_paths:
                        wells=[self._well_by_path(p) for p in self.sequence_paths[len(prefix_paths)-1:]]
                    else:
                        wells=[prefix[-1].result.target]+[b.result.target for b in self.bundles[start_index+1:]]
                    initial=output_markers(prefix[-1].result)
                    following=self.bundles[start_index+1] if start_index+1<len(self.bundles) else prefix[-1]
                    config=dict(following.config)
                    params=Params(**config.get('base_params',asdict(following.result.params)))
                    curves=config.get('requested_curves',following.result.curves)
                if len(wells)<2:
                    raise ValueError('Добавьте минимум две скважины в последовательный набор.')
            else:
                ref,target=self._current()
                if not ref or not target or ref.path==target.path:
                    raise ValueError('Выберите две разные скважины.')
                rm,tm=mapping[ref.path],mapping[target.path]
                if mode=='tune' and (not rm or not tm):
                    raise ValueError('Для подбора нужны маркеры обеих скважин.')
        except (ValueError,TypeError,StopIteration) as error:
            messagebox.showerror('Расчёт',str(error),parent=self)
            return
        def work():
            cb=lambda f,s:self.messages.put(('progress',(f,s)))
            if mode=='sequence':
                bundles,errors=run_sequence(wells,mapping,curves,params,config,self.cancel_event,cb,initial,manual_picks)
                flag_reused_calibration(bundles,calibration_context)
                return 'analysis',dict(bundles=prefix+bundles,errors=errors,mode=mode)
            if config['ensemble'] or mode=='tune':
                bundle=run_ensemble(ref,target,curves,params,rm,tm,config['count'],config['seed'],config['tolerance'],self.cancel_event,cb,mode=='tune')
                bundle.config.update(config)
            else:
                bundle=RunBundle(correlate(ref,target,curves,params,rm,tm,self.cancel_event,cb),config=config)
            if mode!='tune':
                reapply_manual(bundle.result,manual_picks.get(target.path,{}))
                flag_reused_calibration([bundle],calibration_context)
            return 'analysis',dict(bundles=[bundle],errors=[],mode=mode)
        self.status.set('Запуск подбора…' if mode=='tune' else 'Расчёт…')
        self._start_worker(work)

    def _handle_extra(self,kind,payload):
        if kind=='saved':
            self.project_path=payload
            self.dirty=False
            self.status.set(f'Проект сохранён: {payload}')
            return
        if kind=='project':
            path,data=payload
            self._set_busy(False)
            self._restore_project(data)
            self.project_path=path
            self.dirty=False
            self.status.set(f'Проект открыт: {path}')
            return
        self.bundles=payload['bundles']
        self.dirty=True
        if self.bundles:
            self.result=self.bundles[-1].result
            self._show_result()
        else:
            self.result=None
        self.sequence_status.configure(text=f'Рассчитано пар: {len(self.bundles)}' + ('; '+ '; '.join(payload['errors']) if payload['errors'] else ''))
        self._refresh_result_windows()
        if payload['mode']=='tune' and self.bundles:
            self._apply_best(0,show_settings=False)
            self._open_result()
        if payload['errors']:
            messagebox.showwarning('Последовательность остановлена', '\n'.join(payload['errors']),parent=self)

    def _open_result(self):
        if self.bundles:
            window=CorrelationWindow(self)
            self.result_windows.append(window)

    def _refresh_result_windows(self):
        self.result_windows=[w for w in self.result_windows if w.winfo_exists()]
        for window in self.result_windows:
            window.refresh()

    def _manual_changed(self,index):
        self.dirty=True
        result=self.bundles[index].result
        picks=self.manual_picks.setdefault(result.target.path,{})
        picks.update({_key(row['marker']):copy.deepcopy(row) for row in result.rows if row.get('confirmed')})
        for bundle in self.bundles[index+1:]:
            bundle.result.provenance['stale']=True
        self.result=self.bundles[index].result
        self._show_result()
        self._refresh_result_windows()

    def _apply_best(self,index,show_settings=True):
        result=self.bundles[index].result
        if result.provenance.get('mode')=='tuned':
            self.calibration_context[result.target.path]=dict(source=result.ref.path,
                control_names=result.provenance.get('control_names',[]),params=asdict(result.params))
        previous_guard=self._guard
        curve_state=self.curve_list.cget('state')
        self._guard=True
        try:
            for key,value in asdict(result.params).items():
                self.vars[key].set('' if value is None else str(value))
            labels=list(self.wells)
            self.ref_var.set(next(k for k in labels if self.wells[k].path==result.ref.path))
            self.target_var.set(next(k for k in labels if self.wells[k].path==result.target.path))
            self.rgroup_var.set(self.assignments.get(result.ref.path,match_group(result.ref,self.groups)))
            self.tgroup_var.set(self.assignments.get(result.target.path,match_group(result.target,self.groups)))
            # Worker results arrive before the controls are re-enabled. Tk ignores
            # selection changes on a disabled Listbox, so restore its state locally.
            self.curve_list.configure(state='normal')
            self.curve_list.delete(0,'end')
            common=sorted(set(result.ref.curves)&set(result.target.curves))
            for i,name in enumerate(common):
                self.curve_list.insert('end',name)
                if name in result.curves:
                    self.curve_list.selection_set(i)
        finally:
            self.curve_list.configure(state=curve_state)
            self._guard=previous_guard
        self.dirty=True
        self.status.set('Параметры лучшего проверенного варианта применены. Оценка остаётся результатом подбора.')
        if show_settings:
            self._settings()

    def _project_ui(self):
        return dict(params={k:v.get() for k,v in self.vars.items()},config={k:v.get() for k,v in self.config_vars.items()},
                    sequence=self.sequence_paths,assignments=self.assignments,curves=self._selected_curves(),
                    manual_picks=self.manual_picks,
                    calibration_context=self.calibration_context,
                    ref=self._current()[0].path if self._current()[0] else None,
                    target=self._current()[1].path if self._current()[1] else None,
                    marker_path=getattr(self,'marker_path',''))

    def _save_project(self):
        if self.busy:
            return
        path=filedialog.asksaveasfilename(parent=self,title='Сохранить проект',defaultextension='.idtw',
                                         initialfile=Path(self.project_path).name if self.project_path else 'correlation.idtw',filetypes=[('IDTW project','*.idtw')])
        if not path:
            return
        wells,markers,bundles,ui=list(self.wells.values()),list(self.markers),list(self.bundles),self._project_ui()
        def work():
            save_project(path,wells,markers,bundles,ui)
            return 'saved',path
        self._start_worker(work)
        self.cancel_button.configure(state='disabled')

    def _open_project(self):
        if self.busy:
            return
        path=filedialog.askopenfilename(parent=self,title='Открыть проект',filetypes=[('IDTW project','*.idtw')])
        if path:
            self._start_worker(lambda:('project',(path,load_project(path))))

    def _restore_project(self,data):
        wells,markers,bundles,ui=data
        self._guard=True
        self.set_dataset(wells,markers)
        self._guard=True
        for key,value in ui.get('params',{}).items():
            if key in self.vars:
                self.vars[key].set(value)
        for key,value in ui.get('config',{}).items():
            if key in self.config_vars:
                self.config_vars[key].set(value)
        self.assignments=ui.get('assignments',{})
        self.manual_picks=ui.get('manual_picks',{})
        self.calibration_context=ui.get('calibration_context',{})
        for var,key in ((self.ref_var,'ref'),(self.target_var,'target')):
            var.set(next((label for label,w in self.wells.items() if w.path==ui.get(key)),''))
        self._pair_changed()
        self._guard=True
        self.curve_list.selection_clear(0,'end')
        for i in range(self.curve_list.size()):
            if self.curve_list.get(i) in ui.get('curves',[]):
                self.curve_list.selection_set(i)
        self.sequence_paths=[p for p in ui.get('sequence',[]) if p in {w.path for w in wells}]
        self.marker_path=ui.get('marker_path','снимок проекта')
        self.bundles=bundles
        self.result=bundles[-1].result if bundles else None
        self._guard=False
        self._update_files()
        if self.result:
            self._show_result()
        self._refresh_buttons()
        self._refresh_result_windows()

    def _draw(self):
        super()._draw()
        # The detailed result window uses marker-specific colours and ranges.


class CorrelationWindow(tk.Toplevel):
    """Independent result viewport; every plotted object belongs to a snapshot."""
    def __init__(self,app):
        super().__init__(app)
        self.app=app
        self.title('Результат корреляции — набор, варианты, маркеры')
        self.geometry('1320x880')
        self.minsize(900,620)
        self.low=None
        self.high=None
        self._drag=None
        self._track_info=[]
        self.pair_var=tk.StringVar(value='Весь набор')
        self.variant_var=tk.StringVar(value='Итог')
        self.curve_var=tk.StringVar()
        bar=ttk.Frame(self,padding=8)
        bar.pack(fill='x')
        ttk.Label(bar,text='Просмотр:').pack(side='left')
        self.pair_box=ttk.Combobox(bar,textvariable=self.pair_var,state='readonly',width=33)
        self.pair_box.pack(side='left',padx=5)
        self.pair_box.bind('<<ComboboxSelected>>',lambda e:self.refresh(reset=True))
        self.variant_box=ttk.Combobox(bar,textvariable=self.variant_var,state='readonly',width=22)
        self.variant_box.pack(side='left',padx=5)
        self.variant_box.bind('<<ComboboxSelected>>',lambda e:self.refresh(reset=True))
        self.curve_box=ttk.Combobox(bar,textvariable=self.curve_var,state='readonly',width=12)
        self.curve_box.pack(side='left',padx=5)
        self.curve_box.bind('<<ComboboxSelected>>',lambda e:self.draw())
        ttk.Button(bar,text='Весь диапазон',command=self.reset_view).pack(side='left',padx=5)
        ttk.Label(self,text='Колесо — глубина; Shift + колесо — масштаб; перетаскивание — глубина. '
                  'Цвет обозначает маркер; точка — исходный контроль, ромб — ручное подтверждение.').pack(anchor='w',padx=10)
        self.banner=ttk.Label(self,text='',wraplength=1240,foreground='#9b4b08')
        self.banner.pack(fill='x',padx=10,pady=3)
        tabs=ttk.Notebook(self)
        tabs.pack(fill='both',expand=True,padx=8,pady=5)
        graph=ttk.Frame(tabs)
        tableframe=ttk.Frame(tabs)
        details=ttk.Frame(tabs)
        tabs.add(graph,text='Корреляционный профиль')
        tabs.add(tableframe,text='Маркеры и ручные правки')
        tabs.add(details,text='Генерации и метрики')
        self.canvas=tk.Canvas(graph,background='#fafbfd',highlightthickness=0)
        scrollbar=ttk.Scrollbar(graph,orient='horizontal',command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=scrollbar.set)
        scrollbar.pack(side='bottom',fill='x')
        self.canvas.pack(fill='both',expand=True)
        self.canvas.bind('<Configure>',lambda e:self.draw())
        self.canvas.bind('<MouseWheel>',self.wheel)
        self.canvas.bind('<Button-4>',self.wheel)
        self.canvas.bind('<Button-5>',self.wheel)
        self.canvas.bind('<ButtonPress-1>',lambda e:setattr(self,'_drag',(e.y,self.low,self.high)))
        self.canvas.bind('<B1-Motion>',self.pan)
        self.canvas.bind('<ButtonRelease-1>',lambda e:setattr(self,'_drag',None))
        self.canvas.bind('<Motion>',self.hover)
        self.hint=tk.StringVar(value='')
        ttk.Label(self,textvariable=self.hint).pack(fill='x',padx=10,pady=4)
        cols=('pair','marker','predicted_md','original_md','actual_md','error','p10','p90','support','status')
        self.table=ttk.Treeview(tableframe,columns=cols,show='headings',selectmode='browse')
        for key,title in zip(cols,('Пара','Маркер','MD прогноз','MD до правки','MD контроль','Ошибка','P10','P90','Поддержка','Статус')):
            self.table.heading(key,text=title)
            self.table.column(key,width=85 if key not in ('pair','status') else 180,stretch=True)
        ys=ttk.Scrollbar(tableframe,command=self.table.yview)
        xs=ttk.Scrollbar(tableframe,orient='horizontal',command=self.table.xview)
        self.table.configure(yscrollcommand=ys.set,xscrollcommand=xs.set)
        buttons=ttk.Frame(tableframe)
        buttons.pack(side='bottom',fill='x',pady=5)
        ttk.Button(buttons,text='Исправить / подтвердить…',command=self.edit_selected).pack(side='left',padx=4)
        ttk.Button(buttons,text='Пересчитать после этой скважины',command=self.recalculate_after).pack(side='left',padx=4)
        ttk.Button(buttons,text='История правок',command=self.history).pack(side='left',padx=4)
        xs.pack(side='bottom',fill='x')
        ys.pack(side='right',fill='y')
        self.table.pack(fill='both',expand=True)
        self.table.bind('<Double-1>',lambda e:self.edit_selected())
        self.details=tk.Text(details,wrap='word',font=('Consolas',10))
        scroll=ttk.Scrollbar(details,command=self.details.yview)
        self.details.configure(yscrollcommand=scroll.set)
        applybar=ttk.Frame(details)
        applybar.pack(side='bottom',fill='x',pady=4)
        ttk.Button(applybar,text='Применить параметры итогового варианта пары',command=self.apply_best).pack(side='left')
        scroll.pack(side='right',fill='y')
        self.details.pack(fill='both',expand=True)
        self.refresh(reset=True)

    def _pair_index(self):
        value=self.pair_var.get()
        if value=='Весь набор':
            return None
        try:
            index=int(value.split('.',1)[0])-1
            return index if 0<=index<len(self.app.bundles) else None
        except ValueError:
            return None

    def viewed(self):
        index=self._pair_index()
        if index is None:
            return [(i,b.result) for i,b in enumerate(self.app.bundles)]
        bundle=self.app.bundles[index]
        if self.variant_var.get().startswith('Генерация '):
            gen=int(self.variant_var.get().split()[-1])
            result=next((r for r in bundle.candidates if r.provenance.get('generation')==gen),bundle.result)
        else:
            result=bundle.result
        return [(index,result)]

    def refresh(self,reset=False):
        if not self.winfo_exists():
            return
        labels=['Весь набор']+[f'{i+1}. {b.result.ref.name} → {b.result.target.name}' for i,b in enumerate(self.app.bundles)]
        self.pair_box.configure(values=labels)
        if len(labels)==2 and self.pair_var.get()=='Весь набор':
            self.pair_var.set(labels[1])
        if self.pair_var.get() not in labels:
            self.pair_var.set(labels[1] if len(labels)==2 else labels[0])
        index=self._pair_index()
        variants=['Итог']+([f'Генерация {r.provenance["generation"]}' for r in self.app.bundles[index].candidates] if index is not None else [])
        self.variant_box.configure(values=variants)
        if self.variant_var.get() not in variants:
            self.variant_var.set('Итог')
        viewed=self.viewed()
        curves=sorted({c for _,r in viewed for c in r.curves})
        self.curve_box.configure(values=curves)
        if self.curve_var.get() not in curves:
            self.curve_var.set(curves[0] if curves else '')
        for item in self.table.get_children():
            self.table.delete(item)
        warnings=[]
        text=[]
        for i,result in viewed:
            if self.app.bundles[i].result.provenance.get('mode')=='tuned' or self.app.bundles[i].result.provenance.get('calibrated_on_target'):
                warnings.append('Подбор по контрольной разметке: показанная ошибка не является независимой проверкой.')
            if result.provenance.get('stale'):
                warnings.append('Есть результаты, устаревшие после ручной правки. Пересчитайте следующие пары.')
            if result.provenance.get('settings_changed'):
                warnings.append('Показан сохранённый результат прежних настроек; текущие настройки изменены.')
            for j,row in enumerate(result.rows):
                values=[f'{i+1}',row['marker']]+[self.app._fmt(row.get(k)) for k in ('predicted_md','original_md','actual_md','error','p10','p90','support','status')]
                self.table.insert('', 'end',iid=f'{i}:{j}',values=values)
            text.append(f'Пара {i+1}: {result.ref.name} → {result.target.name}\n')
            text.extend(f'{key}: {self.app._fmt(value)}\n' for key,value in result.stats.items())
            text.append('Параметры показанного пути: '+json.dumps(asdict(result.params),ensure_ascii=False)+'\n')
            text.append('Кривые: '+', '.join(result.curves)+'\n')
            for item in self.app.bundles[i].report:
                text.append(f"Генерация {item['generation']}{' [ВЫБРАНА]' if item.get('selected') else ''}: " +
                            json.dumps(_json_clean(item),ensure_ascii=False)+'\n')
            for row in result.rows:
                if row.get('alternatives'):
                    text.append(f"Альтернативы {row['marker']}: "+json.dumps(row['alternatives'],ensure_ascii=False)+'\n')
            text.append('\n')
        self.banner.configure(text=' '.join(dict.fromkeys(warnings)) or 'P10–P90 — чувствительность автоматических генераций внутри группы, до ручных правок. Поддержка — доля генераций, не вероятность.')
        self.details.configure(state='normal')
        self.details.delete('1.0','end')
        self.details.insert('1.0',''.join(text))
        self.details.configure(state='disabled')
        if reset or self.low is None:
            self.reset_view()
        else:
            self.draw()

    def reset_view(self):
        viewed=self.viewed()
        if viewed:
            self.low=min(min(r.zr[0],r.zt[0]) for _,r in viewed)
            self.high=max(max(r.zr[-1],r.zt[-1]) for _,r in viewed)
        else:
            self.low,self.high=0.,100.
        self.draw()

    def wheel(self,event):
        if self.low is None:
            return
        direction=1 if getattr(event,'num',None)==4 or getattr(event,'delta',0)>0 else -1
        span=self.high-self.low
        if getattr(event,'state',0)&1:
            fraction=np.clip((event.y-45)/max(1,self.canvas.winfo_height()-85),0,1)
            anchor=self.low+fraction*span
            span=max(.05,span*(.8 if direction>0 else 1.25))
            self.low=anchor-fraction*span
            self.high=self.low+span
        else:
            self.low-=direction*.1*span
            self.high-=direction*.1*span
        self.draw()

    def pan(self,event):
        if self._drag:
            y,low,high=self._drag
            shift=(event.y-y)*(high-low)/max(1,self.canvas.winfo_height()-85)
            self.low,self.high=low-shift,high-shift
            self.draw()

    def draw(self):
        canvas=self.canvas
        canvas.delete('all')
        viewed=self.viewed()
        h,w=canvas.winfo_height(),canvas.winfo_width()
        if not viewed or h<100 or w<100 or self.low is None:
            return
        top,bottom=50,h-35
        self._track_info=[]
        wells=[viewed[0][1].ref]+[r.target for _,r in viewed]
        bounds=[(viewed[0][1].zr[0],viewed[0][1].zr[-1])]+[(r.zt[0],r.zt[-1]) for _,r in viewed]
        width=max(w,len(wells)*285)
        track_width=width/len(wells)
        canvas.configure(scrollregion=(0,0,width,h))
        def y(md):
            return top+(md-self.low)/(self.high-self.low)*(bottom-top)
        def mark(x0,x1,md,name,kind='predicted',uncertain=False):
            yy=y(md)
            if not top<=yy<=bottom:
                return
            color=marker_color(name)
            canvas.create_line(x0,yy,x1,yy,fill=color,width=2,dash=(4,3) if uncertain else ())
            canvas.create_text(x1-2,yy-3,text=name[:22],anchor='se',fill=color,font=('Segoe UI',9))
            if kind=='manual':
                canvas.create_polygon(x1-6,yy-5,x1-1,yy,x1-6,yy+5,x1-11,yy,fill=color,outline=color)
            elif kind=='control':
                canvas.create_oval(x0-4,yy-4,x0+4,yy+4,fill=color,outline=color)
        curve=self.curve_var.get()
        tracks=[]
        for index,(well,(start,end)) in enumerate(zip(wells,bounds)):
            x0,x1=index*track_width+65,(index+1)*track_width-48
            tracks.append((x0,x1))
            self._track_info.append((x0,x1,well))
            canvas.create_rectangle(x0,top,x1,bottom,fill='white',outline='#b8c5d2')
            canvas.create_text((x0+x1)/2,14,text=well.name[:30],font=('Segoe UI',10,'bold'))
            canvas.create_text((x0+x1)/2,33,text=f'{curve} ({well.units.get(curve, "")}) · MD, м',fill='#526778')
            for md in np.linspace(self.low,self.high,9):
                yy=y(md)
                canvas.create_line(x0,yy,x1,yy,fill='#edf0f4')
                canvas.create_text(x0-5,yy,text=f'{md:.1f}',anchor='e',font=('Segoe UI',9))
            values=well.curves.get(curve)
            if values is not None:
                usable=np.isfinite(values)&(well.depth>=start)&(well.depth<=end)
                visible=usable&(well.depth>=self.low)&(well.depth<=self.high)
                if usable.any():
                    xmin,xmax=float(np.min(values[usable])),float(np.max(values[usable]))
                    scale=max(1e-12,xmax-xmin)
                    indices=np.flatnonzero(visible)
                    breaks=(np.diff(indices)>1)|(np.diff(well.depth[indices])>viewed[min(index,len(viewed)-1)][1].params.gap)
                    for block in np.split(indices,np.flatnonzero(breaks)+1):
                        if len(block)<2:
                            continue
                        take=np.unique(np.r_[block[::max(1,len(block)//1600)],block[-1]])
                        points=[]
                        for sample in take:
                            points.extend((x0+(values[sample]-xmin)/scale*(x1-x0),y(well.depth[sample])))
                        canvas.create_line(*points,fill='#285f91',width=1.1)
                    canvas.create_text(x0,bottom+15,text=f'{xmin:.3g}',anchor='w')
                    canvas.create_text(x1,bottom+15,text=f'{xmax:.3g}',anchor='e')
            else:
                canvas.create_text((x0+x1)/2,(top+bottom)/2,text='Кривая отсутствует',fill='#8a6470')
        for marker in viewed[0][1].reference_markers:
            mark(*tracks[0],marker.md,marker.name,'control')
        for index,(_,result) in enumerate(viewed):
            x0,x1=tracks[index+1]
            for row in result.rows:
                md=row.get('predicted_md')
                if md is None:
                    continue
                color=marker_color(row['marker'])
                if row.get('p10') is not None and row.get('p90') is not None:
                    ya,yb=max(top,y(row['p10'])),min(bottom,y(row['p90']))
                    if ya<=yb:
                        canvas.create_rectangle(x0,ya,x1,yb,fill=color,stipple='gray25',outline='')
                uncertain='провер' in row['status'].lower()
                mark(x0,x1,md,row['marker'],'manual' if row.get('confirmed') else 'predicted',uncertain)
                yr,yt=y(row['ref_md']),y(md)
                if top<=yr<=bottom and top<=yt<=bottom:
                    canvas.create_line(tracks[index][1],yr,x0,yt,fill=color,width=2,dash=(5,3) if uncertain else ())
                for alternative in row.get('alternatives',[]):
                    ya=y(alternative['md'])
                    if top<=ya<=bottom:
                        canvas.create_line(x0,ya,x1,ya,fill=color,dash=(2,5))
                        canvas.create_text(x0+3,ya+3,text=f"альт. {alternative['votes']}/{alternative['total']}",anchor='nw',fill=color,font=('Segoe UI',8))
            for marker in result.control_markers:
                mark(x0,x1,marker.md,marker.name,'control')

    def hover(self,event):
        if self.low is None:
            return
        md=self.low+(event.y-50)/max(1,self.canvas.winfo_height()-85)*(self.high-self.low)
        x=self.canvas.canvasx(event.x)
        track=next(((i,w) for i,(lo,hi,w) in enumerate(self._track_info) if lo<=x<=hi),None)
        if track is None:
            return
        index,well=track
        viewed=self.viewed()
        if index==0:
            result=viewed[0][1]
            ref_md=md
            target_md=float(np.interp(md,result.zr,result.mapped)) if result.zr[0]<=md<=result.zr[-1] else None
        else:
            result=viewed[index-1][1]
            if not result.zt[0]<=md<=result.zt[-1]:
                self.hint.set(f'{well.name} MD {md:.2f} м — вне расчётного интервала')
                return
            j=int(np.argmin(abs(result.mapped-md)))
            ref_md,target_md=float(result.zr[j]),md
        row=min(result.rows,key=lambda r:abs((r['predicted_md'] if r.get('predicted_md') is not None else float('inf'))-(target_md if target_md is not None else 0)),default=None)
        suffix=''
        if row and row.get('predicted_md') is not None and target_md is not None and abs(row['predicted_md']-target_md)<(self.high-self.low)*.03:
            suffix=f" · {row['marker']}: {row['status']}; S={self.app._fmt(row.get('semblance'))}; поддержка={self.app._fmt(row.get('support'))}"
        self.hint.set(f'{well.name} MD {md:.2f} м · соответствие {ref_md:.2f} → {self.app._fmt(target_md)} м'+suffix)

    def _selected_row(self):
        selection=self.table.selection()
        if not selection:
            raise ValueError('Выберите строку маркера.')
        index,row_index=map(int,selection[0].split(':'))
        if self.variant_var.get()!='Итог':
            raise ValueError('Переключитесь на «Итог» для ручного редактирования.')
        return index,row_index,self.app.bundles[index].result.rows[row_index]

    def edit_selected(self):
        if self.app.busy:
            return
        try:
            index,row_index,row=self._selected_row()
        except ValueError as error:
            messagebox.showerror('Маркер',str(error),parent=self)
            return
        dialog=tk.Toplevel(self)
        dialog.title(f"Подтверждение {row['marker']}")
        ttk.Label(dialog,text=f"Исходный автоматический прогноз: {self.app._fmt(row.get('original_md',row['predicted_md']))} м").pack(padx=15,pady=8)
        md=tk.StringVar(value='' if row['predicted_md'] is None else str(row['predicted_md']))
        comment=tk.StringVar()
        ttk.Label(dialog,text='MD в целевой скважине, м:').pack(anchor='w',padx=15)
        ttk.Entry(dialog,textvariable=md,width=35).pack(padx=15,pady=4)
        ttk.Label(dialog,text='Комментарий:').pack(anchor='w',padx=15)
        ttk.Entry(dialog,textvariable=comment,width=60).pack(padx=15,pady=4)
        def accept():
            if self.app.busy:
                return
            try:
                edit_marker(self.app.bundles[index].result,row['marker'],float(md.get().replace(',','.')),comment.get())
                self.app._manual_changed(index)
                dialog.destroy()
            except (ValueError,IndexError) as error:
                messagebox.showerror('Правка',str(error),parent=dialog)
        ttk.Button(dialog,text='Подтвердить',command=accept).pack(pady=12)
        dialog.transient(self)
        dialog.grab_set()

    def recalculate_after(self):
        try:
            index,_,_=self._selected_row()
            self.app._run_mode('sequence',start_index=index)
        except ValueError as error:
            messagebox.showerror('Пересчёт',str(error),parent=self)

    def history(self):
        try:
            _,_,row=self._selected_row()
            messagebox.showinfo('История правок',json.dumps(row.get('history',[]),ensure_ascii=False,indent=2),parent=self)
        except ValueError as error:
            messagebox.showerror('История',str(error),parent=self)

    def apply_best(self):
        if self.app.busy:
            return
        index=self._pair_index()
        if index is None:
            messagebox.showinfo('Параметры','Выберите конкретную пару в верхнем списке.',parent=self)
        else:
            self.app._apply_best(index)


def launch_gui():
    IDTWApp().mainloop()


def gui_smoke():
    """Exercise file selection callbacks, worker queue, drawing and CSV export."""
    import time
    from types import SimpleNamespace
    from openpyxl import Workbook
    ref, target, rm, tm = _synthetic_case()
    app = IDTWApp()
    app.withdraw()
    failures = []
    app.report_callback_exception = lambda kind, error, trace: failures.append(str(error))
    messagebox.showerror = lambda title, text, **kw: failures.append(text)
    messagebox.showwarning = lambda title, text, **kw: failures.append(text)
    def drain():
        deadline = time.monotonic()+30
        while app.busy and time.monotonic()<deadline:
            app.update()
            time.sleep(.01)
        app.update()
        if app.busy or failures:
            raise AssertionError('GUI worker: ' + str(failures or 'timeout'))
    try:
        with tempfile.TemporaryDirectory(prefix='idtw_gui_check_') as directory:
            directory = Path(directory)
            paths = []
            for index, well in enumerate((ref, target)):
                path = directory/f'well_{index}.las'
                path.write_text(_fixture_las(well), encoding='utf-8')
                paths.append(str(path))
            book = Workbook()
            sheet = book.active
            sheet.append(['Маркер','Скважина','UWI','MD','X','Y','Z'])
            for marker in rm+tm:
                sheet.append([marker.name,marker.well,marker.uwi,marker.md,1,2,3])
            workbook_path = directory/'markers.xlsx'
            book.save(workbook_path)
            book.close()
            filedialog.askopenfilenames = lambda **kw: paths
            app._choose_las()
            drain()
            if len(app.wells)!=2 or not app._selected_curves():
                raise AssertionError('LAS selection did not populate pair/curves')
            filedialog.askopenfilename = lambda **kw: str(workbook_path)
            app._choose_markers()
            drain()
            if not app.rgroup_var.get() or not app.tgroup_var.get():
                raise AssertionError('Marker group auto-matching failed')
            app.vars['step'].set('.25')
            app.vars['band'].set('15')
            app._run()
            drain()
            if app.result is None or len(app.table.get_children())!=6:
                raise AssertionError('No result/table after GUI calculation')
            # Offscreen mapping allows real Tk geometry and Canvas layout while
            # avoiding a foreground helper window during an automated check.
            app.geometry('1240x880+20000+20000')
            app.deiconify()
            app.update()
            app._draw()
            if len(app.canvas.find_all()) < 50 or app._plot is None:
                raise AssertionError('Canvas plot was not rendered')
            before = app.result
            app._zoom(SimpleNamespace(y=150,delta=120,state=1))
            app._hover(SimpleNamespace(x=100,y=150))
            if not app._u1-app._u0 < 1 or app.result is not before:
                raise AssertionError('Interactive zoom changed the result')
            app.table.selection_set('0')
            app._select_marker()
            app._reset_view()
            export_path = directory/'result.csv'
            filedialog.asksaveasfilename = lambda **kw: str(export_path)
            app._export()
            with export_path.open(encoding='utf-8-sig',newline='') as stream:
                exported = list(csv.reader(stream,delimiter=';'))
            if len(exported)!=7:
                raise AssertionError('CSV rows missing')
            app.vars['band'].set('20')
            if app.result is not None or str(app.export_button['state'])!='disabled':
                raise AssertionError('Stale result survived input change')
            if failures:
                raise AssertionError(str(failures))
            print('PASS: Tk file loading, background IDTW, marker matching, Canvas, zoom, hover, table, CSV, invalidation.')
    finally:
        app._close()


def advanced_gui_smoke():
    from types import SimpleNamespace
    app=IDTWApp()
    app.withdraw()
    failures=[]
    app.report_callback_exception=lambda kind,error,trace:failures.append(str(error))
    old_error,old_warning=messagebox.showerror,messagebox.showwarning
    messagebox.showerror=lambda title,text,**kw:failures.append(text)
    messagebox.showwarning=lambda title,text,**kw:failures.append(text)
    def drain():
        deadline=time.monotonic()+45
        while app.busy and time.monotonic()<deadline:
            app.update()
            time.sleep(.01)
        app.update()
        if app.busy or failures:
            raise AssertionError(str(failures or 'GUI timeout'))
    try:
        ref,target,rm,tm=_synthetic_case()
        third=Well('third_gui.las','Скважина В','0003',target.depth+15,{'GR':target.curves['GR'].copy()},target.units.copy())
        third_markers=[Marker(m.name,third.name,third.uwi,m.md+15) for m in tm]
        app.set_dataset([ref,target,third],rm+tm+third_markers)
        app._auto_depth()
        app.settings_dialog.withdraw()
        assert app.vars['ref_start'].get()=='0' and app.vars['target_end'].get()=='120'
        for key in ('ref_start','ref_end','target_start','target_end'):
            app.vars[key].set('')
        app.config_vars['count'].set('4')
        app.config_vars['auto'].set(False)
        app.available_list.selection_set(0,2)
        app._sequence_add()
        assert app.sequence_paths==[ref.path,target.path,third.path]
        app.geometry('1280x900+20000+20000')
        app.deiconify()
        app.notebook.select(1)
        app.update()
        box0,box1=app.sequence_list.bbox(0),app.sequence_list.bbox(1)
        app._drag_sequence_start(SimpleNamespace(y=box0[1]+3))
        app._drag_sequence(SimpleNamespace(y=box1[1]+3))
        assert app.sequence_paths[:2]==[target.path,ref.path]
        app._drag_sequence_start(SimpleNamespace(y=box1[1]+3))
        app._drag_sequence(SimpleNamespace(y=box0[1]+3))
        assert app.sequence_paths[:2]==[ref.path,target.path]
        app._run_mode('sequence')
        drain()
        assert len(app.bundles)==2
        app._open_result()
        window=app.result_windows[-1]
        window.geometry('1320x880+20000+20000')
        app.update()
        window.draw()
        assert len(window.canvas.find_all())>60
        span=window.high-window.low
        low=window.low
        window.wheel(SimpleNamespace(y=180,delta=-120,state=0))
        assert abs(window.high-window.low-span)<1e-8 and window.low>low
        window.wheel(SimpleNamespace(y=180,delta=120,state=1))
        assert window.high-window.low<span
        window.reset_view()
        window.table.selection_set('0:0')
        before=app.bundles[0].result.rows[0]['predicted_md']
        window.edit_selected()
        dialog=next(child for child in window.winfo_children() if isinstance(child,tk.Toplevel))
        entries=[w for w in dialog.winfo_children() if isinstance(w,ttk.Entry)]
        entries[0].delete(0,'end')
        entries[0].insert(0,str(before+.2))
        next(w for w in dialog.winfo_children() if isinstance(w,ttk.Button)).invoke()
        assert app.bundles[0].result.rows[0]['confirmed'] and app.bundles[1].result.provenance['stale']
        assert target.path in app.manual_picks
        # Simulate a sequence that stopped before the third well: continue using
        # the requested order, not just the already computed pair list.
        app.bundles=app.bundles[:1]
        app._run_mode('sequence',start_index=0)
        drain()
        assert abs(app.bundles[1].result.reference_markers[0].md-(before+.2))<1e-8
        app._run_mode('tune')
        drain()
        assert app.bundles[0].result.provenance['mode']=='tuned'
        winner=app.bundles[0].result
        assert asdict(app._read_params())==asdict(winner.params)
        assert app._selected_curves()==winner.curves
        assert app.calibration_context[target.path]['params']==asdict(winner.params)
        assert app.result is winner and not winner.provenance.get('settings_changed')
        window=app.result_windows[-1]
        window.pair_var.set(window.pair_box['values'][1])
        window.refresh()
        window.variant_var.set(window.variant_box['values'][1])
        window.refresh()
        assert window.viewed()[0][1] in app.bundles[0].candidates
        window.variant_var.set('Итог')
        window.refresh()
        saved_params={k:v.get() for k,v in app.vars.items()}
        saved_config={k:v.get() for k,v in app.config_vars.items()}
        original_pair=(app.ref_var.get(),app.target_var.get())
        labels=list(app.wells)
        for ref_label,target_label in ((labels[1],labels[2]),(labels[2],labels[0]),original_pair):
            app.ref_var.set(ref_label)
            app.target_var.set(target_label)
            app._pair_changed()
            app.update()
            assert {k:v.get() for k,v in app.vars.items()}==saved_params
            assert {k:v.get() for k,v in app.config_vars.items()}==saved_config
        app.settings_dialog.withdraw()
        app._run_mode('pair')
        drain()
        assert app.bundles[0].result.provenance.get('calibrated_on_target')
        with tempfile.TemporaryDirectory(prefix='idtw_gui_project_') as directory:
            path=str(Path(directory)/'project.idtw')
            filedialog.asksaveasfilename=lambda **kw:path
            app._save_project()
            drain()
            assert Path(path).exists()
            filedialog.askopenfilename=lambda **kw:path
            app._open_project()
            drain()
            assert len(app.wells)==3 and len(app.bundles)==1 and app.manual_picks[target.path]
            assert app.sequence_paths==[ref.path,target.path,third.path]
            assert {k:v.get() for k,v in app.vars.items()}==saved_params
            assert {k:v.get() for k,v in app.config_vars.items()}==saved_config
        if failures:
            raise AssertionError(str(failures))
        print('PASS: advanced GUI auto intervals, drag/drop, three wells, result window, wheel/Shift, manual pick, downstream recalculation, tuning, variant switch, automatic settings, unchanged settings on well switch, project save/load.')
    finally:
        app._close()
        messagebox.showerror,messagebox.showwarning=old_error,old_warning


def main():
    parser = argparse.ArgumentParser(description='IDTW: парная корреляция LAS и перенос маркеров, Tkinter GUI.')
    parser.add_argument('--self-test', action='store_true', help='Проверить алгоритм и форматы входных файлов без окна')
    parser.add_argument('--gui-smoke', action='store_true', help='Проверить Tkinter и сценарий загрузка → расчёт → CSV')
    parser.add_argument('--demo', action='store_true', help='Открыть синтетический пример с известным соответствием')
    args = parser.parse_args()
    if args.self_test:
        self_test()
        advanced_self_test()
    elif args.gui_smoke:
        gui_smoke()
        advanced_gui_smoke()
    elif args.demo:
        ref, target, rm, tm = _synthetic_case()
        app = IDTWApp()
        app.set_dataset([ref,target], rm+tm)
        app.title('IDTW — синтетический демонстрационный пример')
        app.vars['step'].set('.25')
        app.vars['band'].set('15')
        app.after(100,app._run)
        app.mainloop()
    else:
        launch_gui()


if __name__ == '__main__':
    main()
