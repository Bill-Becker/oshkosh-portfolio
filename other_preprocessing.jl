
# Get net metering **availability**, but not limit, by state - could assume 2 MW, 5 MW or unlimited to speed things up
# This came from processing a database which was populated using .csv downloads from DSIRE (update regularly)
# Eaton provided the python script to populate DSIRE.db (SQLite DB) from .csv DSIRE files
incentives_pv = JSON.parsefile("data_states_pv.json")
net_metering_all = Dict([(state, incentives_pv[state]["can_net_meter"]) for state in keys(incentives_pv)])

# Check and update for NEM
# Net metering modeling can be very slow, so instead using export credit = import/retail rate with
# capped/max size based on load, available land/roof, and/or NEM limit
state = df[i, "State"]
net_metering = net_metering_all[state] == 1 ? true : false
if net_metering
    # input_data["ElectricUtility"]["net_metering_limit_kw"] = net_metering_limit_kw
    # input_data_site["PV"]["max_kw"] = input_data["ElectricUtility"]["net_metering_limit_kw"]
    # input_data_site["Wind"]["max_kw"] = input_data["ElectricUtility"]["net_metering_limit_kw"]
    input_data_site["ElectricTariff"]["wholesale_rate"] = electricity_cost_factor * 
                                input_data_site["ElectricTariff"]["blended_annual_energy_rate"] 
else
    input_data_site["ElectricTariff"]["wholesale_rate"] = export_credit_fraction * electricity_cost_factor *
                                input_data_site["ElectricTariff"]["blended_annual_energy_rate"]
end

# Heating load addressable load fraction; paint ovens can't be addressed
addressable_heating_load_fraction = 0.8

# Assumed number of stories for the facility which is used to calculate the facility footprint and subtract from total land
facility_stories = 1


# Land and roof area available for PV and Wind
total_land_acres = df[i, "Facililty Land area (acres)"]
total_facility_sqft = df[i, "Facility square footage"]
facility_footprint_sqft = total_facility_sqft / facility_stories
input_data_site["Site"]["land_acres"] = max(0.0, total_land_acres - facility_footprint_sqft / 43560.0)
input_data_site["Site"]["roof_squarefeet"] = 0.0  #facility_footprint_sqft

# Initially set max size for PV and Wind based on annual load = production estimate, but update to NEM limit if NEM is available
pv_max_load = avg_elec_load_kw / 0.15
pv_max_land = input_data_site["Site"]["land_acres"] / 6.0 * 1000.0
input_data_site["PV"]["max_kw"] = max(0.01,min(pv_max_load, pv_max_land, net_metering_limit_kw))
println("Max PV set to (kW) = ", input_data_site["PV"]["max_kw"])
wind_max_load = avg_elec_load_kw / 0.35
wind_max_land = input_data_site["Site"]["land_acres"] / 30 * 1000.0 # acres per kW * ac, which is only enforced in REopt IF > 1.5MW, so it's sizing to <=1.5MW
input_data_site["Wind"]["max_kw"] = max(0.01,min(wind_max_load, wind_max_land, net_metering_limit_kw))
println("Max Wind set to (kW) = ", input_data_site["Wind"]["max_kw"])
input_data_site["Wind"]["size_class"] = "medium" 

# Calc CHP heuristic that REopt would also calc, but assume 80% boiler effic, 34% elec effic, 44% thermal effic based on recip SC 3
avg_ng_load_mmbtu_per_hour = input_data_site["SpaceHeatingLoad"]["annual_mmbtu"] / 8760 * addressable_heating_load_fraction
chp_heuristic_kw = avg_ng_load_mmbtu_per_hour * 0.8 * 1E6 / 3412 * 1 / 0.34 * 0.44
input_data_site["CHP"]["max_kw"] = min(avg_elec_load_kw, chp_heuristic_kw)  # Rely on max being 2x avg heating load size
# SETTING CHP TO ZERO TO IGNORE FOR THIS CASE
# input_data_site["CHP"]["max_kw"] = 0
# if input_data_site["CHP"]["max_kw"] < 500.0
#     input_data_site["CHP"]["max_kw"] = 0
# end