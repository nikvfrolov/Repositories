# xlsx requires java 64 bit installed
library(comtradr)
library(xlsx)
library(readr)

# Set if there is an "personal library" warning
R_LIBS_SITE="C:\\Temp\\RStudio\\lib"

############ Setup ##################
setwd("C:\\Temp\\RStudio")

# Choose partner country
reporter <- "India"

# Change IP address for each run. UN sets limit on number of queries per hour (VPN)

##########################################

# Partner countries from UN
# Reporter and Partner lists are different!

# On the second run this blocks crashed R with fatal error
# if (!file.exists("partnerAreas.csv")) {
#   download.file("https://unstats.un.org/wiki/download/attachments/79008835/partnerAreas.csv?api=v2",
#                 destfile = "partnerAreas.csv", quiet = TRUE)
# }

# partner_list <- read_csv("partnerAreas.csv", show_col_types = FALSE)
# names(partner_list)[2] <- "Country"

# Unfriendly countries
# http://static.government.ru/media/files/wj1HD7RqdPSxAmDlaisqG2zugWdz8Vc1.pdf

unfriendly <- c("Australia","Albania","Andorra","Iceland","Canada","Japan",
                "South Korea", "North Macedonia", "Liechtenstein","FS Micronesia",
                "Monaco","Montenegro",
                "New Zealand","San Marino","Singapore","Switzerland",
                "Taiwan","Ukraine","United Kingdom","USA","Austria",
                "Belgium","Bulgaria","Croatia","Cyprus","Czechia","Denmark",
                "Estonia","Finland","France","Germany","Greece","Hungary",
                "Ireland","Italy","Latvia","Lithuania","Luxembourg",
                "Malta","Netherlands","Poland","Portugal","Romania",
                "Slovakia","Slovenia","Spain","Sweden","Belgium-Luxembourg")

partner_list$unfriendly <- as.integer(grepl(paste0(unfriendly,collapse = "|"),
                                          partner_list$Country))

# Cleaning partner list
# nes - nothing else specified
noncountries <- c("Africa CAMEU region, nes", "All","World","Areas, nes","Br. Antarctic Terr.",
  "Br. Indian Ocean Terr.","Br. Virgin Isds", "US Virgin Isds","USA (before 1981)",
  "EU-28","CACM, nes","Caribbean, nes","Eastern Europe, nes","Europe EFTA, nes",
  "Europe EU, nes","Fr. South Antarctic Terr.","Free Zones","Holy See (Vatican City State)",
  "India, excl. Sikkim","LAIA, nes","Neutral Zone","North America and Central America, nes",
  "Northern Africa, nes","Oceania, nes","Other Africa, nes","Other Asia, nes",
  "Other Europe, nes","Rest of America, nes","So. African Customs Union",
  "United States Minor Outlying Islands","US Misc. Pacific Isds", "Eswatini")

partner_list$unfriendly[grepl("Fmr",partner_list$Country)] <- 1
partner_list$unfriendly[grepl("Saint",partner_list$Country)] <- 1
partner_list$unfriendly[grepl(paste0(noncountries,collapse = "|"),
                              partner_list$Country)] <- 1

# Split partners list in sets of 5. This is query constraint from UN
partners_we_search <- partner_list$Country[partner_list$unfriendly == 0]
country_list <- split(partners_we_search, ceiling(seq_along(partners_we_search)/5))

for (c in country_list) {
  
  # Comtrade query
  q <- ct_search(reporters = reporter, 
                 partners = c, 
                 trade_direction = c("exports","imports"), 
                 start_date = 2019, 
                 end_date = 2020)
  
  q <- ct_use_pretty_cols(q)
  
    if (exists("all_partners_output")) {
      all_partners_output <- rbind(all_partners_output,q)
    } else {
      all_partners_output <- q
    }
                         }
# Saving results
final_output <- all_partners_output[, c("Year","Trade Flow","Reporter Country",
              "Reporter ISO", "Partner Country","Partner ISO","Trade Value usd")]
write.xlsx(final_output, file = "Trade_table.xlsx", sheetName = reporter, append = TRUE)
