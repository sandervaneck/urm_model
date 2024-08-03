# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 12:48:43 2024

@author: M67B363
"""

import numpy as np
from datetime import date
from ontwikkel import ontwikkel_scenario
import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def get_mediaan_percentile_scenarios(scenarios):
    scenarios = scenarios['results']
    # Extract unique years
    years = set()
    for scenario in scenarios:
        for result in scenario['scenarioresults']:
            years.add(result['jaar'])

    # Prepare a dictionary to store the percentile results for each year
    percentile_results = {year: {'5th': None, '50th': None, '95th': None} for year in years}

    # Process each year separately
    for year in years:
        year_scenarios = []
        for scenario in scenarios:
            for result in scenario['scenarioresults']:
                if result['jaar'] == year:
                    year_scenarios.append({
                        'scenario': scenario['scenario'],
                        'nominal_benefit': result['nominal_benefit'],
                        'savings_honorary': result['capWithContrPostReturn_hon'],
                        'savings_op': result['capWithContrPostReturn'],
                        'real_benefit': result['real_benefit'],
                        'leeftijd': result['leeftijd'],
                        'prognosejaar': result['prognosejaar'],
                        'total_capital': result['total_capital']
                    })
                    break

        # Sort scenarios based on the criteria
        year_scenarios.sort(key=lambda x: (x['nominal_benefit'] > 0, x['nominal_benefit'], x['total_capital']), reverse=False)

        # Calculate the indices for the percentiles
        n = len(year_scenarios)
        if n > 0:
            idx_5th = int(np.ceil(0.05 * n)) - 1
            idx_50th = int(np.ceil(0.50 * n)) - 1
            idx_95th = int(np.ceil(0.95 * n)) - 1

            # Store the results
            percentile_results[year]['5th'] = year_scenarios[max(idx_5th, 0)]
            percentile_results[year]['50th'] = year_scenarios[max(idx_50th, 0)]
            percentile_results[year]['95th'] = year_scenarios[max(idx_95th, 0)]

    return percentile_results

def get_first_year(fourd_filtered, twod_filtered):
    fourd_years = [int(entry['year']) for entry in fourd_filtered if 'year' in entry]
    twod_years = [int(entry['year']) for entry in twod_filtered if 'year' in entry]
    if fourd_years and twod_years:
        prognosejaar = min(min(fourd_years), min(twod_years))
    elif fourd_years:
        prognosejaar = min(fourd_years)
    elif twod_years:
        prognosejaar = min(twod_years)
    else:
        prognosejaar = 1  # Default value if no year found
        
    return prognosejaar
def max_scenario(self):
    scenarios_fourd = [entry['scenario'] for entry in self.fourd]
    scenarios_twod = [entry['scenario'] for entry in self.twod]
    return max(scenarios_fourd + scenarios_twod)

def process_scenario(scenario, nummer, parameters, fourd, twod, endprognoseyear, pensioenleeftijd, pensioenjaar, berekeningsjaar):
    fourd_filtered = [entry for entry in fourd if entry["scenario"] == scenario]
    twod_filtered = [entry for entry in twod if entry["scenario"] == scenario]
    prognosejaar = get_first_year(fourd_filtered, twod_filtered)
    return ontwikkel_scenario(
        endprognoseyear=endprognoseyear,
        scenario=scenario,
        parameters=parameters,
        fourd=fourd_filtered,
        twod=twod_filtered,
        pensioenleeftijd=pensioenleeftijd,
        prognosejaar=prognosejaar,
        pensioenjaar=pensioenjaar,
        eerste_jaar=berekeningsjaar,
        berekeningsjaar=berekeningsjaar
    )

def process_participant(nummer, scenario_aantal, parameters, fourd, twod, endprognoseyear, pensioenleeftijd):
    try:
        if nummer >= len(parameters):
            logging.error(f"Participant index {nummer + 1} out of range for parameters list of length {len(parameters)}")
            return None

        logging.info(f"Processing participant {nummer + 1}...")
        results = []
        for scenario in range(1, scenario_aantal):
            results.append(process_scenario(
                scenario, 
                nummer, 
                parameters[nummer], 
                fourd, 
                twod, 
                endprognoseyear, 
                pensioenleeftijd, 
                pensioenjaar_calculator(parameters[nummer], pensioenleeftijd), 
                parameters[nummer]['calculationdate'].year
            ))
        
        percentile_scenarios = get_percentile_scenarios(results)
        
        pensioenjaar = pensioenjaar_calculator(parameters[nummer], pensioenleeftijd)
        berekeningsjaar = parameters[nummer]['calculationdate'].year
        if pensioenjaar < berekeningsjaar:
            pensioenjaar = berekeningsjaar
        if pensioenjaar >= berekeningsjaar:
            pensioenjaar = min(get_largest_year(percentile_scenarios), pensioenjaar)
        
        mediaan_scenario = percentile_scenarios[pensioenjaar]["50th"]
        porgnosejaar = pensioenjaar - parameters[nummer]['calculationdate'].year + 1
        twods, fourds = get_input_vanaf_prognosejaar(twod, fourd, porgnosejaar)
        pensioenjaar_parameters = creeer_pensioenjaar_parameters(mediaan_scenario, parameters[nummer])
        
        median_results = []
        for scenario in range(1, scenario_aantal):
            median_results.append(process_scenario(
                scenario, 
                nummer, 
                pensioenjaar_parameters, 
                fourds, 
                twods, 
                endprognoseyear, 
                pensioenleeftijd, 
                pensioenjaar, 
                berekeningsjaar
            ))
        
        return {
            "deelnemerresults": {
                "deelnemer": nummer + 1,
                "results": results
            },
            "percentile_results": percentile_scenarios,
            "mediaan_results": {
                "deelnemer": nummer + 1,
                "results": median_results
            }
        }
    except IndexError as e:
        logging.error(f"IndexError for participant {nummer + 1}: {e}")
        return None
    except Exception as e:
        logging.error(f"Unexpected error for participant {nummer + 1}: {e}")
        return None






def get_fourds_prognosis(age, prognosejaar, fourd):
    return {
        "fourd_year0_age0": get_fourd(fourd, age, prognosejaar),
        "fourd_year1_age0": get_fourd(fourd, age, prognosejaar+1),
        "fourd_year0_age1": get_fourd(fourd, age+1, prognosejaar),
        "fourd_year1_age1": get_fourd(fourd, age+1, prognosejaar+1),
        "fourd_year1_age2": get_fourd(fourd, age+2, prognosejaar+1)
        }
def get_fourd(fourd_for_scenario, age, prognosejaar):
    alternative = {
     "year": prognosejaar,
     "scenario": fourd_for_scenario[0]['scenario'],
     "cohort": age,
     "cwf_op": 0.0,
     "cwf_pp": 0.0,
     "total_return": 0.0,
     "total_return_hon": 0.0
     }
    items = [entry for entry in fourd_for_scenario if entry["year"] == prognosejaar and entry["cohort"] == age]
    if len(items) > 0:
        result = items[0]
    else:
        result = alternative
    return result        

def get_twod(twod, prognosejaar):
     return [entry for entry in twod if entry["year"] == prognosejaar][0]
   


def fractie_in_berekeningsjaar_calculator(berekeningsdatum):
    delta = date(berekeningsdatum.year, berekeningsdatum.month, berekeningsdatum.day) - date(berekeningsdatum.year,1,1)
    year_fraction = delta.days / 365.25
    return year_fraction

def get_largest_year(example_variable):
    # Get all keys from the dictionary
    all_keys = example_variable.keys()
    # Find and return the maximum key
    return max(all_keys)

def pensioenjaar_calculator(parameters, pensioenleeftijd):
    result = 0
    if parameters['urm state'] == "retired" or parameters['urm state'] == "partnerPension" or parameters['urm state'] == "orphanPension":
        result = parameters["calculationdate"].year
    else:
        result = min(parameters["default retirement date"].year, parameters["birthdate"].year + pensioenleeftijd)
    return result

def get_input_vanaf_prognosejaar(twods, fourds, prognosejaar):
    twod = [entry for entry in twods if entry["year"] >= prognosejaar]
    fourd = [entry for entry in fourds if entry["year"] >= prognosejaar]
    return (twod, fourd)

def creeer_pensioenjaar_parameters(mediaan_scenario, start_parameters):
    parameters = start_parameters
    parameters['benefit start'] = mediaan_scenario['nominal_benefit']
    parameters['savings start'] = mediaan_scenario['total_capital']
    if parameters['urm state'] != "retired" and parameters['urm state'] != "partnerPension" and parameters['urm state'] != "orphanPension":
       parameters['calculationdate'] = start_parameters['default retirement date']
    parameters['status'] = 'retired'
    return parameters

def get_percentile_scenarios(scenarios):
    # Extract unique years
    years = set()
    for scenario in scenarios:
        for result in scenario['scenarioresults']:
            years.add(result['jaar'])

    # Prepare a dictionary to store the percentile results for each year
    percentile_results = {year: {'5th': None, '50th': None, '95th': None} for year in years}

    # Process each year separately
    for year in years:
        year_scenarios = []
        for scenario in scenarios:
            for result in scenario['scenarioresults']:
                if result['jaar'] == year:
                    year_scenarios.append({
                        'scenario': scenario['scenario'],
                        'nominal_benefit': result['nominal_benefit'],
                        'savings_hon': result['capWithContrPostReturn_hon'],
                        'savings_op': result['capWithContrPostReturn'],
                        'real_benefit': result['real_benefit'],
                        'leeftijd': result['leeftijd'],
                        'prognosejaar': result['prognosejaar'],
                        'total_capital': result['total_capital']
                    })
                    break

        # Sort scenarios based on the criteria
        year_scenarios.sort(key=lambda x: (x['nominal_benefit'] > 0, x['nominal_benefit'], x['total_capital']), reverse=False)

        # Calculate the indices for the percentiles
        n = len(year_scenarios)
        if n > 0:
            idx_5th = int(np.ceil(0.05 * n)) - 1
            idx_50th = int(np.ceil(0.50 * n)) - 1
            idx_95th = int(np.ceil(0.95 * n)) - 1

            # Store the results
            percentile_results[year]['5th'] = year_scenarios[max(idx_5th, 0)]
            percentile_results[year]['50th'] = year_scenarios[max(idx_50th, 0)]
            percentile_results[year]['95th'] = year_scenarios[max(idx_95th, 0)]

    return percentile_results

       