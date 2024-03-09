using REopt
using JuMP
using Xpress
using HiGHS
using JSON
# using REoptPlots
using DataFrames
using CSV
using PyCall
using XLSX
using DotEnv
DotEnv.load!()
# Add a file called .env and add NREL_DEVELOPER_API_KEY="your_api_key_here" on the first line

# Convert city to latitude and longitude
geopy=pyimport("geopy")
geolocator=geopy.geocoders.Nominatim(user_agent="MyApp")

# Import Excel file into DataFrame
df = DataFrame(XLSX.readtable("data.xlsx", "OshKosh Data"))

# Get net metering **availability**, but not limit, by state - could assume 2 MW, 5 MW or unlimited to speed things up
incentives_pv = JSON.parsefile("data_states_pv.json")
net_metering_all = Dict([(state, incentives_pv[state]["can_net_meter"]) for state in keys(incentives_pv)])
net_metering_limit_kw = 5000.0
# For sites which do not have net metering, what is the fraction of retail rate for exported energy
export_credit_fraction = 0.5

# Heating load addressable load fraction; paint ovens can't be addressed
addressable_heating_load_fraction = 0.8

# Assumed number of stories for the facility which is used to calculate the facility footprint and subtract from total land
facility_stories = 2

# Start with a single analysis before setting up loop
input_data = JSON.parsefile("inputs.json")

# # Run REopt for each site after assigning input_data fields specific to the site
# # If not evaluating all sites, for debugging, adjust sites_iter
site_analysis = []
sites_iter = eachindex(df[!, "City"][1:5])
for i in sites_iter
    input_data_site = copy(input_data)
    
    # Site location
    location = geolocator.geocode(df[i, "City"])
    input_data_site["Site"]["latitude"] = location.latitude
    input_data_site["Site"]["longitude"] = location.longitude
    state = df[i, "State"]

    # Land and roof area available for PV and Wind
    total_land_acres = df[i, "Facililty Land area (acres)"]
    total_facility_sqft = df[i, "Facility square footage"]
    facility_footprint_sqft = total_facility_sqft / facility_stories
    input_data_site["Site"]["land_acres"] = total_land_acres - facility_footprint_sqft / 43560.0
    input_data_site["Site"]["roof_squarefeet"] = 0.0  #facility_footprint_sqft

    # Electric load cost
    input_data_site["ElectricLoad"]["annual_kwh"] = df[i, "Annual Electricity, KWH"]
    avg_elec_load_kw = input_data_site["ElectricLoad"]["annual_kwh"] / 8760
    input_data_site["ElectricTariff"]["blended_annual_energy_rate"] = df[i, "\$/kWh"]

    # NG load and cost
    input_data_site["SpaceHeatingLoad"]["annual_mmbtu"] = df[i, "Annual Fuel usage (MMBTU)"]
    input_data_site["SpaceHeatingLoad"]["addressable_load_fraction"] = addressable_heating_load_fraction
    avg_ng_load_mmbtu_per_hour = input_data_site["SpaceHeatingLoad"]["annual_mmbtu"] / 8760 * addressable_heating_load_fraction
    input_data_site["CHP"]["fuel_cost_per_mmbtu"] = df[i, "\$/MMBtu"]
    input_data_site["ExistingBoiler"]["fuel_cost_per_mmbtu"] =  df[i, "\$/MMBtu"]    
    
    # Initially set max size for PV and Wind based on annual load = production estimate, but update to NEM limit if NEM is available
    input_data_site["PV"]["max_kw"] = avg_elec_load_kw / 0.15
    wind_max_load = avg_elec_load_kw / 0.35
    wind_max_land = input_data_site["Site"]["land_acres"] / 0.03 # acres per kW * ac, which is only enforced in REopt IF > 1.5MW, so it's sizing to <=1.5MW
    input_data_site["Wind"]["max_kw"] = min(wind_max_load, wind_max_land)
    input_data_site["Wind"]["size_class"] = "medium" 

    # Check and update for NEM
    net_metering = net_metering_all[state] == 1 ? true : false
    if net_metering
        input_data["ElectricUtility"]["net_metering_limit_kw"] = net_metering_limit_kw
        input_data_site["PV"]["max_kw"] = input_data["ElectricUtility"]["net_metering_limit_kw"]
        input_data_site["Wind"]["max_kw"] = input_data["ElectricUtility"]["net_metering_limit_kw"] 
    else
        input_data_site["ElectricTariff"]["wholesale_rate"] = export_credit_fraction * input_data_site["ElectricTariff"]["blended_annual_energy_rate"]
    end

    # Calc CHP heuristic that REopt would also calc, but assume 80% boiler effic, 34% elec effic, 44% thermal effic based on recip SC 3
    chp_heuristic_kw = avg_ng_load_mmbtu_per_hour * 0.8 * 1E6 / 3412 * 1 / 0.34 * 0.44
    input_data_site["CHP"]["max_kw"] = min(avg_elec_load_kw, chp_heuristic_kw)  # Rely on max being 2x avg heating load size
    if input_data_site["CHP"]["max_kw"] < 500.0
        input_data_site["CHP"]["max_kw"] = 0
    end

    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

    # Xpress solver
    m1 = Model(optimizer_with_attributes(Xpress.Optimizer, "MIPRELSTOP" => 0.01, "OUTPUTLOG" => 0))
    m2 = Model(optimizer_with_attributes(Xpress.Optimizer, "MIPRELSTOP" => 0.01, "OUTPUTLOG" => 0))

    # HiGHS solver
    # m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
    #         "time_limit" => 450.0,
    #         "mip_rel_gap" => opt_tol[i],
    #         "output_flag" => false, 
    #         "log_to_console" => false)
    #         )

    # m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
    #         "time_limit" => 450.0,
    #         "mip_rel_gap" => opt_tol[i],
    #         "output_flag" => false, 
    #         "log_to_console" => false)
    #         )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])
end

# Print tech sizes
# for i in sites_iter
#     for tech in ["PV", "Wind", "CHP"]
#         if haskey(site_analysis[i][2], tech)
#             println("Site $i $tech size (kW) = ", site_analysis[i][2][tech]["size_kw"])
#         end
#     end
# end

# Write results to dataframe

df = DataFrame(site = [i for i in sites_iter], 
               pv_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
               wind_size = [round(site_analysis[i][2]["Wind"]["size_kw"], digits=0) for i in sites_iter],
               chp_size = [round(site_analysis[i][2]["CHP"]["size_kw"], digits=0) for i in sites_iter],
               npv = [round(site_analysis[i][2]["Financial"]["npv"], sigdigits=3) for i in sites_iter],
               renewable_electricity_annual_kwh = [round(site_analysis[i][2]["Site"]["annual_renewable_electricity_kwh"], digits=3) for i in sites_iter],
               emissions_reduction_annual_ton = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"] *
                                                    site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=0) for i in sites_iter]
               )

CSV.write("./portfolio.csv", df)

# results = run_reopt(m1, inputs)

# open("plot_results.json","w") do f
#     JSON.print(f, results)
# end

# plot_electric_dispatch(results)