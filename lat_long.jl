using PyCall
using DataFrames
using JSON
using XLSX

# Import Excel file into DataFrame
df = DataFrame(XLSX.readtable("data.xlsx", "OshKosh Data"))

# Convert city to latitude and longitude
geopy=pyimport("geopy")
geolocator=geopy.geocoders.Nominatim(user_agent="MyApp1")

dict = Dict()
for (i, name) in enumerate(df[!, "Oshkosh Facility Name"])
    # Site location
    city = df[i, "City"]
    dict[city] = Dict()
    # Some cities giving geolocation.geocode(city) trouble; change to nearest working
    if city == "Orlando"
        city = "Clearwater"
    elseif city == "Garner"
        city = "Mason City"
    elseif city == "Bedford"
        city = "Pittsburgh"
    end
    # try
        # TODO get lat/long separately and store in a dict to avoid geolocator calls
    location = geolocator.geocode(city)
    dict[city]["latitude"] = location.latitude
    dict[city]["longitude"] = location.longitude
    # catch
    #     @error("Geolocator errored with city: $city")
    #     continue # skip to next site
    # end
end

open("lat_long.json","w") do f
    JSON.print(f, dict)
end