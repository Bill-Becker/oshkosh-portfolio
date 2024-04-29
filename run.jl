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
using GhpGhx
using DotEnv
DotEnv.load!()

# TODO check that when geocode fails, we're still updating lat/long somehow
# Hybrid GHP with Automatic and Fractional both don't seem to be changing GHX size
# Asked about year one emissions with Cambium to Amanda, but pretty sure we don't get "year one" from Cambium
# Maybe need to use one-year analysis or AVERT to get year one - conscensus with whole team

function set_tech_size!(tech, size, input_data)
    if !(tech == "GHP")
        input_data[tech]["min_kw"] = size
        input_data[tech]["max_kw"] = size
    else
        input_data[tech]["require_ghp_purchase"] = true
    end
    if tech == "CHP"
        input_data[tech]["min_allowable_kw"] = size
    end
end

# Convert city to latitude and longitude
geopy=pyimport("geopy")
geolocator=geopy.geocoders.Nominatim(user_agent="MyApp3")

# Import Excel file into DataFrame
df = DataFrame(XLSX.readtable("data.xlsx", "OshKosh Data"))

# Technologies to evaluate
tech_list = ["GHP"] #, "Wind", "CHP"] #, "GHP"]

# Start with a single analysis before setting up loop, and set all constant inputs
input_data = JSON.parsefile("inputs.json")
input_data["ExistingBoiler"]["production_type"] = "hot_water"
# Electricity cost factor for not being able to reduce demand charges completely
# This is also for fixed charges
electricity_cost_factor = 0.8
# For sites which do not have net metering, what is the fraction of retail rate for exported energy
# this also accounts for NEM which cannot reduce demand charges with exported energy which are baked into the blended rate
export_credit_fraction = 0.5

# Failed runs
failed_runs = []

# For each site, run fixed-size individual techs and write inputs and results to .json file
t = @elapsed begin
    for i in eachindex(df[!, "Oshkosh Facility Name"])
        site_name = df[i, "Oshkosh Facility Name"]
        if df[i, "Eval"] == false
            continue # skip to next site
        end
        # Setup inputs for each scenario (site + tech)
        for tech in tech_list
            filename = site_name * "_" * tech
            input_data_site = copy(input_data)

            # Site location
            city = df[i, "City"]
            # Some cities giving geolocation.geocode(city) trouble; change to nearest working
            if city == "Orlando"
                city = "Clearwater"
            elseif city == "Garner"
                city = "Mason City"
            elseif city == "Bedford"
                city = "Pittsburgh"
            end
            try
                # TODO get lat/long separately and store in a dict to avoid geolocator calls
                location = geolocator.geocode(city)
                input_data_site["Site"]["latitude"] = location.latitude
                input_data_site["Site"]["longitude"] = location.longitude
            catch
                @error("Geolocator errored with city: $city")
                continue # skip to next site
            end

            # Add tech input and set/fix tech size
            input_data_site[tech] = Dict()
            if !(tech == "GHP")
                set_tech_size!(tech, df[i, tech*" Size"], input_data_site)
            else
                input_data_site[tech]["require_ghp_purchase"] = true
            end

            # Roof-mounted PV
            if tech == "PV"
                input_data_site["PV"]["array_type"] = 1
            end

            # Required CHP input
            if tech == "CHP"
                input_data_site["CHP"]["fuel_cost_per_mmbtu"] = df[i, "\$/MMBtu"]
            end

            # Required building_sqft input for GHP, but ignore O&M benefit
            if tech == "GHP"
                input_data_site[tech]["building_sqft"] = df[i, "Facility square footage"]
                input_data_site[tech]["om_cost_per_sqft_year"] = 0.0  # Default is -$0.51/sqft
                aux_heater_type = "electric"
                input_data_site[tech]["is_ghx_hybrid"] = true
                input_data_site[tech]["aux_heater_installed_cost_per_mmbtu_per_hr"] = 0.0
                input_data_site[tech]["aux_cooler_installed_cost_per_ton"] = 0.00
                input_data_site[tech]["ghpghx_inputs"] = [Dict()]
                input_data_site[tech]["ghpghx_inputs"][1]["hybrid_ghx_sizing_method"] = "Fractional" # "Automatic"
                input_data_site[tech]["ghpghx_inputs"][1]["hybrid_ghx_sizing_fraction"] = 0.6
            end

            # Electric load cost
            input_data_site["ElectricLoad"]["annual_kwh"] = df[i, "Annual Electricity, KWH"]
            input_data_site["ElectricTariff"]["blended_annual_energy_rate"] = df[i, "\$/kWh"] * electricity_cost_factor

            # Cooling Load, for GHP analysis
            # input_data_site["CoolingLoad"]["annual_fraction_of_electric_load"] = df[i, "Cooling Load Fraction"]
            input_data_site["CoolingLoad"]["annual_tonhour"] = df[i, "Cooling Load Annual Ton-Hr"]

            # NG load and cost
            input_data_site["SpaceHeatingLoad"]["annual_mmbtu"] = df[i, "Addressable Fuel Load MMBtu"]
            input_data_site["ExistingBoiler"]["fuel_cost_per_mmbtu"] =  df[i, "\$/MMBtu"]

            # Set profile based on operating shifts, mostly weekdays whether 2 or 3 shifts
            if df[i, "Operating Shifts"] == 2
                baseload = "FlatLoad_16_5"
            else
                baseload = "FlatLoad_24_5"
            end
            input_data_site["ElectricLoad"]["blended_doe_reference_names"] = ["Warehouse", baseload]
            input_data_site["ElectricLoad"]["blended_doe_reference_percents"] = [0.5, 0.5]
            input_data_site["SpaceHeatingLoad"]["blended_doe_reference_names"] = ["Warehouse", baseload]
            input_data_site["SpaceHeatingLoad"]["blended_doe_reference_percents"] = [0.5, 0.5]
            input_data_site["CoolingLoad"]["blended_doe_reference_names"] = ["Warehouse", baseload]
            input_data_site["CoolingLoad"]["blended_doe_reference_percents"] = [0.5, 0.5]

            # Check and update for NEM, but avoiding modeling NEM (slow) because setting PV/Wind max_kw to NEM limit
            # Note, the electricity_cost_factor is already applied to the blended_annual_energy_rate to account for lack of reducing demand charges
            if df[i, "Net Metering Limit"] > 0.0
                input_data_site["ElectricTariff"]["wholesale_rate"] = input_data_site["ElectricTariff"]["blended_annual_energy_rate"] 
            else
                input_data_site["ElectricTariff"]["wholesale_rate"] = (export_credit_fraction *
                                                                    input_data_site["ElectricTariff"]["blended_annual_energy_rate"])
            end

            # Run REopt
            results = Dict()
            try
                s = Scenario(input_data_site)
                inputs = REoptInputs(s)

                # Xpress solver
                m1 = Model(optimizer_with_attributes(Xpress.Optimizer, "MIPRELSTOP" => 0.01, "OUTPUTLOG" => 0))
                m2 = Model(optimizer_with_attributes(Xpress.Optimizer, "MIPRELSTOP" => 0.01, "OUTPUTLOG" => 0))

                # HiGHS solver
                # m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
                #         "time_limit" => 450.0,
                #         "mip_rel_gap" => 0.01,
                #         "output_flag" => false, 
                #         "log_to_console" => false)
                #         )

                # m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
                #         "time_limit" => 450.0,
                #         "mip_rel_gap" => 0.01,
                #         "output_flag" => false, 
                #         "log_to_console" => false)
                #         )            

                results = run_reopt([m1,m2], inputs)
            catch
                println("Something went wrong running REopt")
                append!(failed_runs, [filename])
            end

            analysis_data = Dict("inputs" => input_data_site,
                                "outputs" => results)

            filename *= "hybrid"
            open("results/$filename.json","w") do f
                JSON.print(f, analysis_data)
            end

            if !(results["status"] == "optimal")
                append!(failed_runs, filename)
            end
        end
    end
end


println("Time (sec) = ", t)