# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 14:13:58 2024

@author: M67B363
"""
from datetime import date
from parsers import convert_to_float
import math

def get_twod(twod, prognosejaar):
     return [entry for entry in twod if entry["year"] == prognosejaar][0]
   
def cwf_calculator(xi, tau, ppRatio, age, fourd_year0_age0, fourd_year0_age1, fourd_year1_age1, fourd_year1_age2):
    
    cwf_op_j0 = (1-xi) * fourd_year0_age0['cwf_op'] + xi * fourd_year0_age1['cwf_op']
    cwf_pp_j0 = (1-xi) * fourd_year0_age0['cwf_pp'] + xi * fourd_year0_age1['cwf_pp']
    
    cwf_op_j1 = (1-xi) * fourd_year1_age1['cwf_op'] + xi * fourd_year1_age2['cwf_op']
    cwf_pp_j1 = (1-xi) * fourd_year1_age1['cwf_pp'] + xi * fourd_year1_age2['cwf_pp']
    
    cwf_op = (1-tau) * cwf_op_j0 + tau * cwf_op_j1
    cwf_pp = (1-tau) * cwf_pp_j0 + tau * cwf_pp_j1
    
    return cwf_op + ppRatio * cwf_pp

def calculate_nominal_benefit(retiring, nominal_benefit, pensioenjaar, total_capital, ff, twod_year0, jaar, ppRatio, xi, tau, age, fourd_year0_age0, fourd_year0_age1, fourd_year1_age1, fourd_year1_age2):
    
    if pensioenjaar == jaar and retiring:
        cwf = cwf_calculator(
            ppRatio = ppRatio, 
            xi = xi,
            tau = tau,
            age = age,
            fourd_year0_age0 = fourd_year0_age0,
            fourd_year0_age1 = fourd_year0_age1,
            fourd_year1_age1 = fourd_year1_age1,
            fourd_year1_age2 = fourd_year1_age2   
            )
        nominal_benefit = total_capital / (cwf*ff)
    else:
        nominal_benefit = nominal_benefit * (1+twod_year0['payout_adjustment'])
    return nominal_benefit
def tau_voor_jaar(pensioenjaar, jaar, parameters):
    if pensioenjaar == jaar:
        tau = convert_to_float(parameters['year fraction retirement year'])
    elif pensioenjaar > jaar:
        tau = 0
    else: 
        tau = 1
    return tau

def update_retiring(age, pensioenleeftijd, results, i, updatedprognosejaar, parameters, xi):
    retiring = parameters["urm state"] == "retired"
    if math.ceil(age + xi) == pensioenleeftijd:
        retiring = True
    return retiring    

def update_leeftijd(startage_nonrounded, updatedprognosejaar, i, results):
    if results[i] is not None and results[i]['prognosejaar'] == updatedprognosejaar: 
        leeftijd = math.ceil(startage_nonrounded - 1 + updatedprognosejaar)
    else:
        leeftijd = results[i]['leeftijd'] + 1
    return leeftijd

def update_jaar(results, i, updatedprognosejaar):
    if results[i]['prognosejaar'] == updatedprognosejaar:
        jaar = results[i]['jaar']
    else:
        jaar = results[i]['jaar'] + 1
    return jaar    
    
def ontwikkel_startjaar(prognosejaar, ppRatio, fractie_in_berekeningsjaar, xi, pensioenjaar, jaar, age, parameters, twod_year0, twod_year1, fourd_year0_age0, fourd_year1_age0, fourd_year0_age1, fourd_year1_age1, fourd_year1_age2, pensioenleeftijd):
    tau = tau_voor_jaar(pensioenjaar, jaar, parameters)
    fractie_pensioenjaar = parameters['year fraction retirement year']
    pensionBase = parameters['pensionbase'] * twod_year0['cpi']
    contribution = pensionBase * twod_year0["contribution_rate"]
    capWithContr = parameters['savings start'] + contribution if age <= pensioenleeftijd else 0
        
    ret_op = get_return(
        retiring = False, age = age, pensioenleeftijd = pensioenleeftijd,
        returns_age1=fourd_year0_age1['total_return'], returns=fourd_year0_age0['total_return'], xi=xi, 
        fractie_in_pensioenjaar=fractie_pensioenjaar)
    ret_hon = get_return(
        retiring = False, age = age, pensioenleeftijd = pensioenleeftijd,
        returns_age1=fourd_year0_age1['total_return_hon'], returns=fourd_year0_age0['total_return_hon'], xi=xi, 
        fractie_in_pensioenjaar=fractie_pensioenjaar)

    capWithContrPostReturn = capWithContr * ret_op if age <= pensioenleeftijd else 0
    capWithContrPostReturn_hon = parameters['savings start honorary'] * ret_hon if age <= pensioenleeftijd else 0
    total_capital = capWithContrPostReturn + capWithContrPostReturn_hon 
    ff = (1-tau)*twod_year0['ff'] + tau * twod_year1['ff'] 
        
    nominal_benefit = parameters['benefit start']
    nominal_benefit = calculate_nominal_benefit(
        retiring = False,
        nominal_benefit=nominal_benefit,
        pensioenjaar = pensioenjaar, 
        total_capital = total_capital, 
        ff = ff, 
        twod_year0 = twod_year0, 
        jaar = jaar, 
        ppRatio = ppRatio, 
        xi = xi, 
        tau = tau, 
        age = age, 
        fourd_year0_age0 = fourd_year0_age0,
        fourd_year0_age1 = fourd_year0_age1,
        fourd_year1_age1 = fourd_year1_age1,
        fourd_year1_age2 = fourd_year1_age2, 
        )        
    
    nominal_benefit_sr = nominal_benefit * (1+twod_year0["sr_adjustment"]) if jaar >= pensioenjaar else 0
    survivorspension = nominal_benefit_sr * ppRatio
    real_benefit = nominal_benefit_sr / (twod_year0['cpi'] * (1+twod_year0['one_year_inflation'])**fractie_pensioenjaar) if jaar >= pensioenjaar else 0
    
    retired = pensioenjaar <=jaar
    savings_op = 0 if retired else capWithContrPostReturn - nominal_benefit
    savings_hon = 0 if retired else capWithContrPostReturn_hon - survivorspension/(1+twod_year0["sr_adjustment"])
    
    return {
        "jaar": jaar,
        "prognosejaar": prognosejaar,
        "leeftijd": age,
        "tau": tau,
        "nominal_benefit": nominal_benefit,
        "nominal_benefit_sr": nominal_benefit_sr,
        "real_benefit": real_benefit,
        "capWithContrPostReturn": capWithContrPostReturn,
        "capWithContrPostReturn_hon": capWithContrPostReturn_hon,
        "survivorspension": survivorspension,
        "total_capital": total_capital,
        "savings_op": savings_op,
        "savings_hon": savings_hon
        }
    
def get_return(returns, returns_age1, xi, fractie_in_pensioenjaar, retiring, age, pensioenleeftijd):
    result = 1
    a = 1 + returns
    b = 1 + returns_age1
    condition1 = retiring
    condition2 = age != pensioenleeftijd
    if (condition1): 
        if (condition2):
            result = a**fractie_in_pensioenjaar
        else:
            result = (a**(1-xi))*(b**(fractie_in_pensioenjaar-(1-xi)))
    else:
        result = (a**(1-xi))*(b**xi)
    return result

def ontwikkel_jaar(previous_prognosejaar, retiring, previous_tau, nominal_benefit, savings_op, savings_hon, prognosejaar, ppRatio, fractie_in_berekeningsjaar, xi, pensioenjaar, jaar, age, parameters, twod_year0, twod_year1, fourd_year0_age0, fourd_year1_age0, fourd_year0_age1, fourd_year1_age1, fourd_year1_age2, pensioenleeftijd):
    
    tau = tau_voor_jaar(pensioenjaar, jaar, parameters)
    fractie_pensioenjaar = parameters['year fraction retirement year']
    pensionBase = parameters['pensionbase'] * twod_year0['cpi']
    contribution = pensionBase * twod_year0["contribution_rate"] 
    if tau > 0 and tau < 1:
        contribution = contribution * tau
    elif tau == 1:
        contribution = 0
        
    capWithContr = savings_op + contribution if previous_tau == 0 else 0
    
    next_year_same_prognose_year = previous_prognosejaar + parameters['calculationdate'].year == pensioenjaar and parameters['urm state'] != "retired" and previous_prognosejaar != prognosejaar
    returns_1 = fourd_year1_age0
        
    if next_year_same_prognose_year:
        returns_1=fourd_year0_age0
   
    ret_op = get_return(
        retiring = retiring, age = age, pensioenleeftijd = pensioenleeftijd,
        returns_age1=fourd_year0_age1['total_return'], returns=fourd_year0_age0['total_return'], xi=xi, 
        fractie_in_pensioenjaar=fractie_pensioenjaar)
    ret_hon = get_return(
        retiring = retiring, age = age, pensioenleeftijd = pensioenleeftijd,
        returns_age1=fourd_year0_age1['total_return_hon'], returns=fourd_year0_age0['total_return_hon'], 
        xi=xi, fractie_in_pensioenjaar=fractie_pensioenjaar)
    
    capWithContrPostReturn = capWithContr * ret_op 
    capWithContrPostReturn_hon = savings_hon * ret_hon if previous_tau == 0 else 0
    total_capital = capWithContrPostReturn + capWithContrPostReturn_hon 
    ff = (1-tau)*twod_year0['ff'] + tau * twod_year1['ff'] 
        
    nominal_benefit = calculate_nominal_benefit(
        retiring = retiring,
        nominal_benefit=nominal_benefit,
        pensioenjaar = pensioenjaar, 
        total_capital = total_capital, 
        ff = ff, 
        twod_year0 = twod_year0, 
        jaar = jaar, 
        ppRatio = ppRatio, 
        xi = xi, 
        tau = tau, 
        age = age, 
        fourd_year0_age0 = fourd_year0_age0,
        fourd_year0_age1 = fourd_year0_age1,
        fourd_year1_age1 = fourd_year1_age1,
        fourd_year1_age2 = fourd_year1_age2, 
        )
        
    retired = pensioenjaar <=jaar and retiring == False
    nominal_benefit_sr = nominal_benefit * (1+twod_year0["sr_adjustment"]) if jaar >= pensioenjaar else 0
    survivorspension = nominal_benefit_sr * ppRatio
    real_benefit = nominal_benefit_sr / (twod_year0['cpi'] * (1+twod_year0['one_year_inflation'])**fractie_pensioenjaar) if jaar >= pensioenjaar else 0
    savings_op = 0 if retired else capWithContrPostReturn - nominal_benefit
    
    savings_hon = 0 if retired else capWithContrPostReturn_hon - survivorspension/(1+twod_year0["sr_adjustment"])
    return {
        "jaar": jaar,
        "prognosejaar": prognosejaar,
        "leeftijd": age,
        "tau": tau,
        "nominal_benefit": nominal_benefit,
        "nominal_benefit_sr": nominal_benefit_sr,
        "real_benefit": real_benefit,
        "capWithContrPostReturn": capWithContrPostReturn,
        "capWithContrPostReturn_hon": capWithContrPostReturn_hon,
        "total_capital": total_capital,
        "survivors_pension": survivorspension,
        "savings_op": savings_op,
        "savings_hon": savings_hon
        }

def fractie_in_berekeningsjaar_calculator(berekeningsdatum):
    delta = date(berekeningsdatum.year, berekeningsdatum.month, berekeningsdatum.day) - date(berekeningsdatum.year,1,1)
    year_fraction = delta.days / 365.25
    return year_fraction
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
def ontwikkel_scenario(eerste_berekeningsjaar, endprognoseyear, scenario, parameters, twod, fourd, pensioenleeftijd, prognosejaar, pensioenjaar, eerste_jaar, berekeningsjaar):
    
    xi = convert_to_float(parameters['age fraction'])
    ppRatio = convert_to_float(parameters['pp ratio'])
    fractie_in_bjaar = fractie_in_berekeningsjaar_calculator(parameters['calculationdate'])
    startage_nonrounded = (parameters['calculationdate'] - parameters['birthdate']).days / 365.25
    startage = math.floor(startage_nonrounded)
    
    prognosejaar = prognosejaar
    twod_year0 = get_twod(twod, prognosejaar)
    twod_year1 = get_twod(twod, prognosejaar + 1)
    fourd_prognosis = get_fourds_prognosis(age=startage, prognosejaar=prognosejaar, fourd=fourd)
     
    results = []
    jaar = eerste_jaar
    t0 = ontwikkel_startjaar(
        prognosejaar = prognosejaar, 
        ppRatio = ppRatio, 
        fractie_in_berekeningsjaar = fractie_in_bjaar, 
        xi = xi,
        pensioenjaar = pensioenjaar,
        jaar = jaar, 
        age = startage, 
        parameters = parameters,
        twod_year0 = twod_year0, 
        twod_year1 = twod_year1,
        fourd_year0_age0 = fourd_prognosis["fourd_year0_age0"],
        fourd_year1_age0 = fourd_prognosis["fourd_year1_age0"],
        fourd_year0_age1 = fourd_prognosis["fourd_year0_age1"],
        fourd_year1_age1 = fourd_prognosis["fourd_year1_age1"],
        fourd_year1_age2 = fourd_prognosis["fourd_year1_age2"],  
        pensioenleeftijd = pensioenleeftijd
        )
    
    results.append(t0)
    laatste_kalender_jaar = pensioenjaar + 11
    i=0
    age = startage
    
    while prognosejaar + eerste_berekeningsjaar < min(endprognoseyear, laatste_kalender_jaar):
        previous_prognosejaar = prognosejaar
        prognosejaar += 1
        age += 1
        jaar += 1
        retiring = parameters["urm state"] != "retired" and math.ceil(age + xi) == pensioenleeftijd

        twod_year0 = get_twod(twod, prognosejaar)
        fourd_prognosis = get_fourds_prognosis(age=age, prognosejaar=prognosejaar, fourd=fourd)
        twod_year1 = get_twod(twod, prognosejaar + 1)
        result_for_year = ontwikkel_jaar(
            previous_prognosejaar=previous_prognosejaar,
            retiring=retiring,
            previous_tau=results[i]['tau'],
            nominal_benefit= results[i]['nominal_benefit'],
            savings_op = results[i]["savings_op"],
            savings_hon = results[i]["savings_hon"],
            prognosejaar = prognosejaar,
            ppRatio = ppRatio, 
            fractie_in_berekeningsjaar = fractie_in_bjaar, 
            xi = xi,
            pensioenjaar = pensioenjaar,
            jaar = jaar, 
            age = age, 
            parameters = parameters,
            twod_year0 = twod_year0, 
            twod_year1 = twod_year1,
            fourd_year0_age0 = fourd_prognosis["fourd_year0_age0"],
            fourd_year1_age0 = fourd_prognosis["fourd_year1_age0"],
            fourd_year0_age1 = fourd_prognosis["fourd_year0_age1"],
            fourd_year1_age1 = fourd_prognosis["fourd_year1_age1"],
            fourd_year1_age2 = fourd_prognosis["fourd_year1_age2"],      
            pensioenleeftijd = pensioenleeftijd
            )
        #print(f"berekening prognosejaar {prognosejaar} af")
        i += 1
        results.append(result_for_year)      
    
    return {
        "scenario": scenario,
        "status": parameters["urm state"],
        "op_te_bouwen_jaren": pensioenjaar-berekeningsjaar + 1,
        "scenarioresults": results
        }