
using DataFrames
using JSON
using XLSX

# This file creates a results_summary.json file with all site-technology scenario data
# TODO make site-level summary which aggregates tech-scenario metrics like NPV and CO2_Reduced_Tonne

# Import Excel file into DataFrame
df = DataFrame(XLSX.readtable("data.xlsx", "OshKosh Data"))
tech_list = ["PV", "Wind", "CHP"] #, "GHP"]

# Incentives used for calcs
itc_fraction = 0.3
iac_grant = 300.0E3

results_summary_dict = Dict()
for (i, site) in enumerate(df[!, "Oshkosh Facility Name"])
    if df[i, "Eval"] == false
        continue
    end
    results_summary_dict[site] = Dict()
    for tech in tech_list
        filename = site * "_" * tech
        # TODO if this fails, continue on but mark data as 0's or NaN or something
        try
            analysis = JSON.parsefile("results/$filename.json")        
            results_summary_dict[site][tech] = Dict()
            if tech == "GHP"
                size_s = analysis["outputs"][tech]["size_heat_pump_ton"]
            else
                size_s = analysis["outputs"][tech]["size_kw"]
            end
            npv_s = round(analysis["outputs"]["Financial"]["npv"], sigdigits=3) + iac_grant
            capex_before_incent = analysis["outputs"]["Financial"]["initial_capital_costs"]
            if capex_before_incent > 2*iac_grant
                capex_s = round(capex_before_incent * (1 - itc_fraction) - iac_grant, sigdigits=4)
            else
                capex_s = round(capex_before_incent * (1 - itc_fraction), sigdigits=4)
            end
            if capex_s > 0.0
                npvi_s = round(npv_s / capex_s, digits=2)
            else
                npvi_s = 0.0
            end
            elec_savings = (analysis["outputs"]["ElectricTariff"]["year_one_bill_before_tax_bau"] - 
                            analysis["outputs"]["ElectricTariff"]["year_one_bill_before_tax"])
            fuel_savings = (analysis["outputs"]["ExistingBoiler"]["year_one_fuel_cost_before_tax_bau"] - 
                            analysis["outputs"]["ExistingBoiler"]["year_one_fuel_cost_before_tax"])
            
            if tech == "CHP"
                fuel_savings -= analysis["outputs"]["CHP"]["year_one_fuel_cost_before_tax"] 
            end
            om_cost = analysis["outputs"]["Financial"]["year_one_om_costs_before_tax"]
            year_one_savings = elec_savings + fuel_savings - om_cost
            spp_s = round(max(0.0, capex_s / year_one_savings), digits=2)
            if isnan(spp_s) || isnothing(spp_s) || isinf(spp_s)
                spp_s = 0.0
            end
            energy_produced_mwh_s = round(analysis["outputs"]["Site"]["annual_renewable_electricity_kwh"] / 1E3, digits=0)
            if tech == "CHP"
                energy_produced_mwh_s += round(analysis["outputs"]["CHP"]["annual_electric_production_kwh"] / 1E3, digits=0)
                energy_produced_mwh_s += round(analysis["outputs"]["CHP"]["annual_thermal_production_mmbtu"] * 1E3 / 3412.0, digits=0)
            end
            annual_co2_reduced_tonne_s = round(analysis["outputs"]["Site"]["annual_emissions_tonnes_CO2"], digits=0)
            # if haskey(analysis["outputs"]["Financial"], "breakeven_cost_of_emissions_reduction_per_tonne_CO2")
            #     co2_breakeven_cost_per_tonne = round(analysis["outputs"]["Financial"]["breakeven_cost_of_emissions_reduction_per_tonne_CO2"], digits=0)
            # else
            #     co2_breakeven_cost_per_tonne = 0.0
            # end
            results_summary_dict[site][tech]["Size"] = round(size_s, digits=0)
            results_summary_dict[site][tech]["NPV"] = npv_s
            results_summary_dict[site][tech]["CapEx"] = capex_s
            results_summary_dict[site][tech]["NPVI"] = npvi_s
            results_summary_dict[site][tech]["Simple_Payback"] = spp_s
            results_summary_dict[site][tech]["Energy_Produced_MWh"] = energy_produced_mwh_s
            results_summary_dict[site][tech]["CO2_Reduced_Year_One_Tonne"] = annual_co2_reduced_tonne_s  # Year one because we just use the first year emissions
            # results_summary_dict[site][tech]["CO2_Breakeven_Cost_Per_Tonne"] = co2_breakeven_cost_per_tonne
        catch
            @warn("Assigning zeros for $site for $tech because did not properly run")
            results_summary_dict[site][tech]["Size"] = 0.0
            results_summary_dict[site][tech]["NPV"] = 0.0
            results_summary_dict[site][tech]["CapEx"] = 0.0
            results_summary_dict[site][tech]["NPVI"] = 0.0
            results_summary_dict[site][tech]["Simple_Payback"] = 0.0
            results_summary_dict[site][tech]["Energy_Produced_MWh"] = 0.0
            results_summary_dict[site][tech]["CO2_Reduced_Year_One_Tonne"] = 0.0            
        end
    end
end

open("results_summary.json","w") do f
    JSON.print(f, results_summary_dict)
end

