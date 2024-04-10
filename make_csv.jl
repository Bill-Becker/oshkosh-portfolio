using JSON
using DataFrames
using CSV
using XLSX


results = JSON.parsefile("results_summary_Pierce -Assembly.json")

# This method of dataframe creation provides control of order of metrics, but at the cost of needing to 
#   sync with metrics in the results summary data file
# df = DataFrame(Site_Tech = site_tech,
#     Size = zeros(length(site_tech)),
#     NPV = zeros(length(site_tech)),
#     CapEx = zeros(length(site_tech)),
#     NPVI = zeros(length(site_tech)),
#     Simple_Payback = zeros(length(site_tech)),
#     RE_Produced_kWh = zeros(length(site_tech)),
#     CO2_Reduced_Tonne = zeros(length(site_tech)),
#     CO2_Breakeven_Cost_Per_Tonne = zeros(length(site_tech))
#     )

# Find a single site and tech to index the results dict for keys
global site1 = ""
global tech1 = ""
for site in keys(results)
    global site1 = site
    for tech in keys(results[site])
        global tech1 = tech
        break
    end
    break
end

# Initialize the dataframe of results metrics
site_tech = [site*"_"*tech for site in keys(results) for tech in keys(results[site])]
df = DataFrame(Site_Tech = site_tech)
for metric in keys(results[site1][tech1])
    df[!, metric] = zeros(length(site_tech))
end

global i = 0
for (s, site) in enumerate(keys(results))
    for (t, tech) in enumerate(keys(results[site]))
        global i += 1
        if !(results[site][tech] == Dict())
            for metric in keys(results[site][tech])
                df[i, metric] = round(results[site][tech][metric], sigdigits=3)
            end
        end
    end
end

CSV.write("results_summary_Pierce.csv", df)