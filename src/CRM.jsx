import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';

const RAW_CONTACTS = [{"id": "2752729795", "first": "Joey", "last": "Cantrell", "title": "General Manager, Operations Central Service Center Maintenance Machine & Roll Shop Facility", "company": "constellium.com", "phone": "(256) 577-2972", "email": "", "city": "Tuscumbia", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/joey-cantrell-cmrt-crl-mmcp-90a01a48"}, {"id": "3875830961", "first": "Shawn", "last": "Pope", "title": "Maintenance Operations Manager, Facilities", "company": "walmart.com", "phone": "(704) 838-6541", "email": "", "city": "Ocean Isle Beach", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/sean-pope-16423a37"}, {"id": "2355701044", "first": "Jason", "last": "Robinson", "title": "Region Maintenance Manager, Facilities", "company": "walmart.com", "phone": "", "email": "", "city": "East Chattanooga", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/jason-robinson-6125ba133"}, {"id": "-2036615052", "first": "Richard", "last": "DuMonte", "title": "General Manager, Facilities", "company": "aramark.com", "phone": "(281) 658-8726", "email": "", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/rick-dumonte-8b01644"}, {"id": "2536974631", "first": "Stephanie", "last": "Martinez", "title": "General Manager I, Facility", "company": "cleanharbors.com", "phone": "(219) 777-1916", "email": "", "city": "Bartow", "state": "Florida", "linkedin": ""}, {"id": "2369839651", "first": "Ralph", "last": "Leon", "title": "Senior Manager Ll, Regional Facilities Manager.", "company": "walmart.com", "phone": "(305) 213-7692", "email": "", "city": "Charlotte", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/ralph-leon-8a2983a1"}, {"id": "1160042904", "first": "Eric", "last": "Benge", "title": "Senior Manager II, Supply Chain Fleet Maintenance", "company": "walmart.com", "phone": "(479) 277-9855", "email": "", "city": "East Fayetteville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/eric-benge-b6a34617a"}, {"id": "16793913606", "first": "Tami", "last": "King", "title": "Manager I, Property Tax", "company": "walmart.com", "phone": "", "email": "", "city": "East Fayetteville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/tami-king-2a6905179"}, {"id": "1235151028", "first": "David", "last": "Case", "title": "Divisional Maintenance Manager", "company": "walmart.com", "phone": "(307) 633-5050", "email": "", "city": "Gainesville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/david-case-4b49345b"}, {"id": "15777475605", "first": "Robert", "last": "John", "title": "Area Maintenance Manager", "company": "walmart.com", "phone": "", "email": "", "city": "Orlando", "state": "Florida", "linkedin": "https://www.linkedin.com/in/bob-john-973110203"}, {"id": "8722899885", "first": "Colton", "last": "Chasteen", "title": "Mechanical Engineer & Project Manager (Capital )", "company": "internationalpaper.com", "phone": "(256) 365-6266", "email": "", "city": "Anniston", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/coltonchasteen"}, {"id": "1645973892", "first": "Roger", "last": "Cooke", "title": "Project Manager ( Construction )", "company": "gd.com", "phone": "(336) 698-8627", "email": "", "city": "Greensboro", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/roger-g-cooke-9a466124"}, {"id": "1726358456", "first": "Andrew", "last": "Lichtenthal", "title": "Manager, Maintenance", "company": "landolakesinc.com", "phone": "(615) 315-4078", "email": "", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/andy-lichtenthal-20250813"}, {"id": "9488616859", "first": "Tom", "last": "Ryder", "title": "E13 Maintenance Manager", "company": "constellium.com", "phone": "(256) 620-0359", "email": "", "city": "Moulton", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/tom-ryder-93306713"}, {"id": "11342770310", "first": "Bill", "last": "Dudley", "title": "Director, Facilities Management", "company": "aramark.com", "phone": "(901) 857-4528", "email": "", "city": "Goodlettsville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/bill-dudley-02563810"}, {"id": "1655989644", "first": "Robert", "last": "Biamonte", "title": "Manager, Facilities", "company": "corporate.comcast.com", "phone": "(386) 847-7070", "email": "", "city": "Fort Lauderdale", "state": "Florida", "linkedin": "https://www.linkedin.com/in/bob-biamonte-5a7a8a16"}, {"id": "1780511852", "first": "Jennifer", "last": "Shroads", "title": "Manager, Engineering & Maintenance", "company": "landolakesinc.com", "phone": "(607) 341-5584", "email": "", "city": "Boiling Springs", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/tom-whitehouse-995b9a47"}, {"id": "8081920784", "first": "Sheldon", "last": "McWilliams", "title": "Manager I, Warehouse Facility", "company": "mohawkind.com", "phone": "(980) 987-1847", "email": "", "city": "Charlotte", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/sheldon-mcwilliams-43803649"}, {"id": "1886492784", "first": "Jodi", "last": "Bell", "title": "Manager, Engineering, Maintenance & Information Technology", "company": "adient.com", "phone": "(615) 424-9748", "email": "", "city": "Murfreesboro", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/jodi-bell-14679a14"}, {"id": "-1871302937", "first": "Paul", "last": "Jones", "title": "Manager, Maintenance", "company": "kochfoods.com", "phone": "(423) 923-5530", "email": "", "city": "Whitesburg", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/paul-jones-64a7266b"}, {"id": "2998444167", "first": "Brian", "last": "Carter", "title": "Project Manager, Facilities", "company": "bridgestoneamericas.com", "phone": "(615) 438-1778", "email": "", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/brian-carter-114485a"}, {"id": "1941462266", "first": "Leslie", "last": "McDougald", "title": "Manager, Maintenance", "company": "ball.com", "phone": "(803) 459-7437", "email": "", "city": "Bishopville", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/leslie-mcdougald-742b2844"}, {"id": "2729930719", "first": "Andrew", "last": "Allen", "title": "Manager, Facilities", "company": "packagingcorp.com", "phone": "(919) 280-2401", "email": "", "city": "Concord", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/andrew-allen-b3b10b32"}, {"id": "-1876598590", "first": "James", "last": "Merager", "title": "Manager, Facility Environmental", "company": "gp.com", "phone": "(803) 960-0742", "email": "", "city": "Sumter", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/jim-merager-b3aa9630"}, {"id": "-2041994703", "first": "Dennis", "last": "Schnurbusch", "title": "Manager Building Maintenance", "company": "bridgestoneamericas.com", "phone": "", "email": "", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/dennis-schnurbusch-3b42206a"}, {"id": "6020321508", "first": "Super", "last": "Scott", "title": "Manager, Maintenance", "company": "invenergy.com", "phone": "", "email": "", "city": "Parrish", "state": "Florida", "linkedin": "https://www.linkedin.com/in/super-scott-231659151"}, {"id": "2649951565", "first": "Kevin", "last": "Kersh", "title": "Senior Manager - Maintenance - Shift", "company": "geappliancesco.com", "phone": "(256) 552-9547", "email": "", "city": "Decatur", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/kevin-kersh-bb9b5b11"}, {"id": "2075913046", "first": "Jesse", "last": "Gillam", "title": "Senior Manager, Facilities", "company": "bridgestoneamericas.com", "phone": "(615) 708-7084", "email": "", "city": "Hendersonville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/jesse-gillam-55736b7a"}, {"id": "2173592031", "first": "Kara", "last": "Greer", "title": "Manager, Facilities & Environmental", "company": "bridgestoneamericas.com", "phone": "(615) 440-0283", "email": "", "city": "Lavergne", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/karagreer"}, {"id": "5908803256", "first": "Chandler", "last": "Hall", "title": "Manager, Facility Operations", "company": "pepsico.com", "phone": "(904) 652-3963", "email": "", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/chandler-hall-9a0a94159"}, {"id": "1212340249", "first": "Michael", "last": "Veihl", "title": "Manager, Maintenance Engineering", "company": "bridgestoneamericas.com", "phone": "(919) 603-5220", "email": "", "city": "Timberlake", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/michaeljveihl"}, {"id": "2001118789", "first": "Clark", "last": "Greene", "title": "Manager, Maintenance", "company": "republicservices.com", "phone": "(803) 818-3837", "email": "", "city": "Yow Mill", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/clark-greene-390b2713"}, {"id": "1909485461", "first": "Neil", "last": "Brogden", "title": "Manager, Distribution & Maintenance", "company": "aramark.com", "phone": "(919) 282-2291", "email": "", "city": "Oxford", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/neil-brogden-04781577"}, {"id": "1291469458", "first": "Jonathan", "last": "Sharp", "title": "Manager, Facilities", "company": "aramark.com", "phone": "(610) 662-0491", "email": "", "city": "Simpsonville", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/john-sharp-4103863b"}, {"id": "3338704267", "first": "Lisa", "last": "Shertza", "title": "Manager, Maintenance", "company": "pepsico.com", "phone": "(813) 975-4406", "email": "", "city": "Saint Augustine", "state": "Florida", "linkedin": "https://www.linkedin.com/in/lisa-shertza-39131a215"}, {"id": "8668897926", "first": "Jason", "last": "Brown", "title": "Manager, Facilities", "company": "aramark.com", "phone": "(856) 676-2521", "email": "", "city": "Gainesville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jason-brown-47106b188"}, {"id": "-2039123776", "first": "Voiers", "last": "Richard", "title": "Manager, Facility", "company": "aramark.com", "phone": "", "email": "", "city": "Morganton", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/voiers-richard-16410a60"}, {"id": "1964316578", "first": "Bill", "last": "Covington", "title": "Manager, Maintenance", "company": "ball.com", "phone": "(615) 596-1868", "email": "", "city": "Dickson", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/bill-covington-ab6aa84b"}, {"id": "9337219020", "first": "Jasmine", "last": "Stringer", "title": "Manager, A2l Facilities", "company": "aramark.com", "phone": "(662) 351-4576", "email": "", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jasminestringer112a"}, {"id": "9437595075", "first": "louis-stephane", "last": "yannantuono", "title": "Manager, A2l Facilities", "company": "aramark.com", "phone": "(360) 949-1414", "email": "louis-stephaneyannantuono@zoomhubs.com", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jasminestringer112a"}, {"id": "11599677393", "first": "Dominic", "last": "Pacheco", "title": "Plant Engineer & Manager, Maintenance", "company": "amcor.com", "phone": "(540) 205-0550", "email": "", "city": "Columbia", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/dominic-pacheco-89847866"}, {"id": "11607163401", "first": "Toral", "last": "Patel", "title": "Manager, Property", "company": "7-eleven.com", "phone": "", "email": "", "city": "Jupiter", "state": "Florida", "linkedin": "https://www.linkedin.com/in/toral-patel-444917265"}, {"id": "10564962105", "first": "Joe", "last": "Sampson", "title": "Manager, Maintenance", "company": "amcor.com", "phone": "(864) 400-3336", "email": "", "city": "Blythewood", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/joe-sampson-3870b295"}, {"id": "2594360340", "first": "Mark", "last": "Erikzon", "title": "Manager, Maintenance", "company": "amcor.com", "phone": "(704) 340-4835", "email": "", "city": "Charlotte", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/mark-erikzon-85144187"}, {"id": "14807738640", "first": "Michael", "last": "Crittenden", "title": "Manager, Site Maintenance & Reliability", "company": "gp.com", "phone": "", "email": "", "city": "Muscle Shoals", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/michael-crittenden-344bba60"}, {"id": "11254351763", "first": "Dale", "last": "Harris", "title": "Manager, Maintenance", "company": "gp.com", "phone": "(850) 593-6151", "email": "", "city": "Sneads", "state": "Florida", "linkedin": "https://www.linkedin.com/in/dale-harris-83a78710a"}, {"id": "2371335254", "first": "Bruce", "last": "Buckoski", "title": "Project Manager, Site Facilities", "company": "novartis.com", "phone": "(919) 793-6035", "email": "", "city": "Chapel Hill", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/bruce-buckoski-43a21511"}, {"id": "3373207250", "first": "Jesse", "last": "Ramos", "title": "Manager, Maintenance", "company": "aramark.com", "phone": "(407) 973-7438", "email": "", "city": "Merritt Island", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jesse-ramos-32084a155"}, {"id": "13970507880", "first": "Dan", "last": "Gunter", "title": "Area Maintenance Manager", "company": "internationalpaper.com", "phone": "(256) 548-2227", "email": "", "city": "Scottsboro", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/dan-gunter-691741283"}, {"id": "9724818114", "first": "David", "last": "Bohannon", "title": "Maintenance Technician & Manager, Line", "company": "amcor.com", "phone": "", "email": "", "city": "Ocoee", "state": "Florida", "linkedin": "https://www.linkedin.com/in/david-bohannon-318128231"}, {"id": "15384808060", "first": "Robert", "last": "Jackson", "title": "Manager, Maintenance", "company": "ppg.com", "phone": "(636) 288-4434", "email": "", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/robertjacksonj1123"}, {"id": "5247151629", "first": "Allen", "last": "Jean", "title": "Manager, Maintenance", "company": "aramark.com", "phone": "", "email": "", "city": "Goodlettsville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/allen-jean-228696a4"}, {"id": "7273377239", "first": "Ed", "last": "Brown", "title": "Manager, Facilities", "company": "abb.com", "phone": "(662) 416-4689", "email": "", "city": "Selmer", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/ed-brown-026b1b204"}, {"id": "1947749073", "first": "Laurie", "last": "Inman", "title": "Construction & Project Management Financial Analyst", "company": "aramark.com", "phone": "(615) 761-0389", "email": "", "city": "Goodlettsville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/ACwAAAtBH5YBeHgZaXsa5e9RhThH7HqCLGPZqpM"}, {"id": "-1241500421", "first": "Kenneth", "last": "Creek", "title": "Maintenance Management Coordinator", "company": "generalmills.com", "phone": "(615) 225-1350", "email": "", "city": "Murfreesboro", "state": "Tennessee", "linkedin": ""}, {"id": "3264148669", "first": "Timothy", "last": "Chandler", "title": "Manager, Facility Capital, Riverdale Mill", "company": "internationalpaper.com", "phone": "(912) 238-7850", "email": "timothy.chandler@ipaper.com", "city": "Montgomery", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/tchand"}, {"id": "9677784416", "first": "Travis", "last": "Greene", "title": "Manager, Maintenance", "company": "nucor.com", "phone": "(903) 626-6217", "email": "greene@vulcraft.com", "city": "Clemmons", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/travis-greene-5ba1847b"}, {"id": "4059040135", "first": "Tristan", "last": "Lewis", "title": "Project Manager, Construction (US)", "company": "7-eleven.com", "phone": "(571) 353-4219", "email": "tlewis@7-eleven.com", "city": "Tampa", "state": "Florida", "linkedin": "https://www.linkedin.com/in/tristan-lewis-b39b07202"}, {"id": "1708437781", "first": "Genice", "last": "Jackson", "title": "Manager, Capital Projects", "company": "internationalpaper.com", "phone": "(251) 937-1904", "email": "genice.jackson@ipaper.com", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/genice-jackson-mba-2a364127"}, {"id": "5984595547", "first": "Steven", "last": "Conner", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "", "email": "steve.conner@walmart.com", "city": "Forest City", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/steve-conner-9b0722126"}, {"id": "3289255930", "first": "Gregory", "last": "Abadie", "title": "Manager, Manufacturing Maintenance Department", "company": "pepsico.com", "phone": "(601) 502-3918", "email": "gregory.abadie@pepsico.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/greg-abadie-1301a3130"}, {"id": "9486636519", "first": "Forrest", "last": "Butts", "title": "Maintenance Set-up Planner J3 General Manager TII", "company": "internationalpaper.com", "phone": "(850) 937-5096", "email": "forrest_butts@internationalpaper.com", "city": "Orange Park", "state": "Florida", "linkedin": ""}, {"id": "-2027578655", "first": "James", "last": "Large", "title": "Sheet Finishing Maintenance Manager", "company": "constellium.com", "phone": "(256) 386-6608", "email": "james.large@constellium.com", "city": "Muscle Shoals", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/james-large-8a2189108"}, {"id": "6001091774", "first": "John", "last": "Hatfield", "title": "General Manager, Maintenance Technology J3", "company": "internationalpaper.com", "phone": "(901) 262-7077", "email": "john.hatfield@ipaper.com", "city": "Brewton", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/john-hatfield-6223b415"}, {"id": "1810615108", "first": "Johnathan", "last": "Odom", "title": "Manager, Engineering & Maintenance", "company": "smurfitwestrock.com", "phone": "(843) 430-5329", "email": "john.odom@westrock.com", "city": "Florence", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/john-odom-50925963"}, {"id": "9862922383", "first": "Felix", "last": "Mezquita", "title": "Manager, Maintenance", "company": "cornerstonebuildingbrands.com", "phone": "(713) 705-9152", "email": "felix.mezquita@cornerstone-bb.com", "city": "Cary", "state": "North Carolina", "linkedin": ""}, {"id": "2364836433", "first": "Dwayne", "last": "McCullion", "title": "Project Manager, Large Capital", "company": "gp.com", "phone": "(361) 242-8632", "email": "dwayne.mccullion@gapac.com", "city": "Seneca", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/dwayne-mccullion"}, {"id": "1887377211", "first": "Paul G", "last": "Rasmussen", "title": "General Manager II, Facility Wwt (Chattanooga)", "company": "cleanharbors.com", "phone": "(423) 825-4114", "email": "rasmussen.paul@cleanharbors.com", "city": "Chattanooga", "state": "Tennessee", "linkedin": ""}, {"id": "9954077102", "first": "Dallas", "last": "Gabel", "title": "Senior Manager, Operations & Facility Improvement", "company": "walmart.com", "phone": "(909) 801-0107", "email": "dallas.gabel@walmart.com", "city": "Apollo Beach", "state": "Florida", "linkedin": "https://www.linkedin.com/in/dallasgabel"}, {"id": "2786576231", "first": "Jeffrey", "last": "Dekker", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "(252) 430-5650", "email": "jeffrey.dekker@walmart.com", "city": "Henderson", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/jef-dekker-2b67701bb"}, {"id": "5453573546", "first": "Kenneth", "last": "Mathis", "title": "Manager, Maintenance", "company": "kochfoods.com", "phone": "(334) 387-1732", "email": "kenneth.mathis@kochfoods.com", "city": "Montgomery", "state": "Alabama", "linkedin": ""}, {"id": "2660857956", "first": "Anthony", "last": "Palazzolo", "title": "Manager, Capex Project & Facilities Maintenance", "company": "pepsico.com", "phone": "(412) 849-9649", "email": "anthony.palazzolo@pepsico.com", "city": "Charlotte Harbor", "state": "Florida", "linkedin": "https://www.linkedin.com/in/anthony-palazzolo-62a10352"}, {"id": "-2048641582", "first": "Robert", "last": "Burdette", "title": "Project Manager, Maintenance", "company": "honeywell.com", "phone": "(864) 444-6227", "email": "rburdette@honeywell.com", "city": "Isle of Palms", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/robert-burdette-a1b5895a"}, {"id": "1907508455", "first": "Debbie", "last": "Greathouse", "title": "Facilities Management Financial Analyst", "company": "aramark.com", "phone": "(615) 761-0363", "email": "greathouse-debbie@aramark.com", "city": "Goodlettsville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/debbie-greathouse-21b96a115"}, {"id": "-1319888228", "first": "Daryl", "last": "Henderson", "title": "Manager, Maintenance", "company": "kraton.com", "phone": "(850) 464-0711", "email": "daryl.henderson@kraton.com", "city": "Cantonment", "state": "Florida", "linkedin": "https://www.linkedin.com/in/daryl-henderson-34a554106"}, {"id": "-2040459712", "first": "Jeffrey", "last": "Welsch", "title": "Manager, Maintenance", "company": "owenscorning.com", "phone": "(904) 549-2148", "email": "jeffrey.a.welsch@owenscorning.com", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jeffrey-welsch-cmrp-crl-753b5b72"}, {"id": "9882059898", "first": "Darnell", "last": "Stephens", "title": "Manager, Maintenance", "company": "mauserpackaging.com", "phone": "(336) 637-1408", "email": "darnell.stephens@mauserpackaging.com", "city": "Burlington", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/darnell-stephens-0a2237282"}, {"id": "-862586740", "first": "James", "last": "Stasiak", "title": "Program Manager, Predictive Maintenance Rcm Coe", "company": "adm.com", "phone": "(224) 637-9437", "email": "james.stasiak@adm.com", "city": "Cary", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/joey-stasiak"}, {"id": "8256172004", "first": "Michael", "last": "Norris", "title": "Manager, Maintenance", "company": "owenscorning.com", "phone": "(864) 296-4185", "email": "mnorris@owenscorning.com", "city": "Anderson", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/michael-norris-475b45222"}, {"id": "1925070865", "first": "Carolina", "last": "Aguilar", "title": "Senior Manager, Sds Purchases Facilities & Eps", "company": "us.pg.com", "phone": "+506 2204 5270", "email": "aguilar.ci@pg.com", "city": "Gainesville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/carolina-aguilar-371766127"}, {"id": "8145131121", "first": "Frank", "last": "Hanson", "title": "Manager, Facility", "company": "smithfieldfoods.com", "phone": "(910) 308-3259", "email": "frankhanson@smithfield.com", "city": "Shannon", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/frank-hanson-b58a058a"}, {"id": "3293487491", "first": "Adam", "last": "Young", "title": "Senior Manager, Indirects Procurement Facilities & Real Estate", "company": "kraftheinzcompany.com", "phone": "", "email": "adam.young1@kraftheinz.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/adam-young-210420b3"}, {"id": "5458250497", "first": "Jayson", "last": "Robinson", "title": "General Manager, Maintenance Technology Ia", "company": "internationalpaper.com", "phone": "(850) 937-5012", "email": "jason.robinson@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": ""}, {"id": "3184787980", "first": "Christopher", "last": "Zindell", "title": "Manager, Maintenance", "company": "ppg.com", "phone": "(336) 477-8372", "email": "czindell@ppg.com", "city": "Lexington", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/christopher-zindell-58aa5946"}, {"id": "13633374887", "first": "Kevin", "last": "Shyrer", "title": "Manager, Maintenance", "company": "novolex.com", "phone": "(904) 783-9985", "email": "kevin.shyrer@novolex.com", "city": "", "state": "Florida", "linkedin": "https://www.linkedin.com/in/kevin-shyrer-532768237"}, {"id": "2550384217", "first": "Aglathia", "last": "Thompson-Hicks", "title": "Manager, Area Facilities", "company": "7-eleven.com", "phone": "(806) 336-5914", "email": "ahicks@7-eleven.com", "city": "Orlando", "state": "Florida", "linkedin": "https://www.linkedin.com/in/aglathia-thompson-hicks-88525413"}, {"id": "3251694630", "first": "Marc", "last": "Hughes Jr", "title": "General Manager, Maintenance Technology J3", "company": "internationalpaper.com", "phone": "(850) 937-4506", "email": "marc.hughes@ipaper.com", "city": "Memphis", "state": "Tennessee", "linkedin": ""}, {"id": "1831625402", "first": "Brandon", "last": "Keeney", "title": "Director, Integrated Facilities Management", "company": "aramark.com", "phone": "(217) 898-5384", "email": "keeney-brandon@aramark.com", "city": "Greenville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/brandon-keeney-cmrp-pmp-4012ba38"}, {"id": "5981321810", "first": "Emily", "last": "Dearden", "title": "Manager, Total Productive Maintenance At Pepsi (Tpm)", "company": "pepsico.com", "phone": "", "email": "emily.dearden@pepsico.com", "city": "Jacksonville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/emily-dearden-023067184"}, {"id": "2833202096", "first": "Michael", "last": "Wooten", "title": "Manager, Facilities Preservation", "company": "internationalpaper.com", "phone": "(803) 481-6700", "email": "michael.wooten@ipaper.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/michael-wooten-95271682"}, {"id": "14516885080", "first": "Hilary", "last": "Brandau", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(334) 782-0407", "email": "hilary.brandau@ipaper.com", "city": "Prattville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/hilary-walterscheid"}, {"id": "2020070137", "first": "Paul", "last": "Howell", "title": "Manager, Facilities & Engineering", "company": "goldenstatefoods.com", "phone": "(334) 610-3103", "email": "pahowell@goldenstatefoods.com", "city": "Opelika", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/paul-howell-aa605340"}, {"id": "-1785620623", "first": "Jackie", "last": "Hucks", "title": "Manager, Maintenance", "company": "domtar.com", "phone": "(910) 610-5266", "email": "jackie.hucks@domtar.com", "city": "", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/jackie-hucks-1590549b"}, {"id": "5336587109", "first": "Thomas", "last": "Watson", "title": "Manager, Maintenance", "company": "novolex.com", "phone": "(423) 826-1734", "email": "thomas.watson@novolex.com", "city": "Chattanooga", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/thomas-watson-02b861b6"}, {"id": "1580533782", "first": "Paul", "last": "Quinley", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(251) 937-1900", "email": "paul.quinley@ipaper.com", "city": "", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/paul-quinley-2546ab39"}, {"id": "11291885286", "first": "Robert", "last": "Beasley", "title": "Manager, Maintenance", "company": "gaf.com", "phone": "(803) 984-3098", "email": "rbeasley@gaf.com", "city": "Lancaster", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/robert-beasley-cmrp-0376227b"}, {"id": "-1879288021", "first": "Smith", "last": "Bobby", "title": "Manager, Maintenance", "company": "republicservices.com", "phone": "(912) 585-8045", "email": "smithbobby@republicservices.com", "city": "West Palm Beach", "state": "Florida", "linkedin": "https://www.linkedin.com/in/smith-bobby-a0b50848"}, {"id": "7394986501", "first": "Roy", "last": "Roderick", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "(863) 517-9672", "email": "roy.roderick@walmart.com", "city": "Fort Lauderdale", "state": "Florida", "linkedin": "https://www.linkedin.com/in/roy-roderick-iii-a66815172"}, {"id": "8475518508", "first": "Brian", "last": "Jeffords", "title": "Project Manager, Capital", "company": "internationalpaper.com", "phone": "(843) 545-2495", "email": "brian.jeffords@bfusa.com", "city": "Pawleys Island", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/brian-jeffords-0244b8b8"}, {"id": "2989322879", "first": "Brian", "last": "Brown", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(407) 484-8498", "email": "brown.brian@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": "https://www.linkedin.com/in/brian-brown-78357135"}, {"id": "931904523", "first": "Jerry", "last": "Becker", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(615) 896-2240", "email": "jerrybecker@ipaper.com", "city": "Murfreesboro", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/jerry-becker-49706028"}, {"id": "11717270860", "first": "Andrew", "last": "Pron", "title": "General Manager, Facilities", "company": "aramark.com", "phone": "(267) 760-3373", "email": "pron-andrew@aramark.com", "city": "Orlando", "state": "Florida", "linkedin": "https://www.linkedin.com/in/andypron"}, {"id": "3078282197", "first": "Robert", "last": "Ciaramitaro", "title": "Manager, Industrial Maintenance Operations", "company": "walmart.com", "phone": "(772) 467-4500", "email": "robert.ciaramitaro@walmart.com", "city": "Fort Pierce", "state": "Florida", "linkedin": "https://www.linkedin.com/in/robert-ciaramitaro-a067378b"}, {"id": "2839198201", "first": "Matthew", "last": "Thein", "title": "Project Manager, Facilities", "company": "bridgestoneamericas.com", "phone": "(217) 971-0143", "email": "theinmatthew@bfusa.com", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/matthewthein"}, {"id": "16457540306", "first": "James", "last": "Martin", "title": "Director, Facilities Management, P3 Schools, Pgmd", "company": "honeywell.com", "phone": "(410) 910-0860", "email": "jim.r.martin@honeywell.com", "city": "Charlotte", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/james-martin-fma%c2%ae-rpa%c2%ae-5856a31a"}, {"id": "8080007002", "first": "Walter", "last": "Clark", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "(803) 277-5640", "email": "wclark@walmart.com", "city": "Greenville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/walter-clark-a5938914"}, {"id": "9126773361", "first": "Lee", "last": "Gallman", "title": "Manager, Maintenance", "company": "kraftheinzcompany.com", "phone": "(803) 944-9702", "email": "doyle.gallman@kraftheinz.com", "city": "Prosperity", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/lee-gallman-199538174"}, {"id": "14133645131", "first": "Charles", "last": "Chasteen", "title": "Manager, Section -Facilities", "company": "gd.com", "phone": "(850) 766-5999", "email": "char@gd.com", "city": "Tallahassee", "state": "Florida", "linkedin": "https://www.linkedin.com/in/charles-chasteen-a7091246"}, {"id": "14882433018", "first": "Luciano", "last": "Santana Ferreira", "title": "Manager, Maintenance", "company": "gaf.com", "phone": "", "email": "luciano.santanaferreira@gaf.com", "city": "Tampa", "state": "Florida", "linkedin": "https://www.linkedin.com/in/luciano-santana-ferreira-547985380"}, {"id": "12173538995", "first": "Thomas", "last": "Conry", "title": "Fleet Maintenance Manager", "company": "republicservices.com", "phone": "(734) 355-4564", "email": "tconry@republicservices.com", "city": "Lebanon", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/thomas-conry-91842b1b5"}, {"id": "5405568125", "first": "Mark", "last": "Erikzon", "title": "Manager, Maintenance", "company": "mauserpackaging.com", "phone": "(704) 393-6430", "email": "mark.erikzon@mauserpackaging.com", "city": "Charlotte", "state": "North Carolina", "linkedin": ""}, {"id": "6032989351", "first": "Andrew", "last": "Johnson", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(434) 522-2737", "email": "andy.johnson@ipaper.com", "city": "Lynchburg", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/andy-johnson-08958b11b"}, {"id": "1964024500", "first": "Mark", "last": "Koumas", "title": "Manager, Area Maintenance Finished Products & Customer Products Services (1 2 Paper Machines)", "company": "internationalpaper.com", "phone": "(803) 353-7628", "email": "mark.koumas@ipaper.com", "city": "Eastover", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/mark-koumas-32700733"}, {"id": "5328410430", "first": "Randall", "last": "Putnam", "title": "Manager, Corporate Facilities", "company": "internationalpaper.com", "phone": "(901) 419-7236", "email": "randall.putnam@ipaper.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/randall-putnam-3673474"}, {"id": "1644810210", "first": "Carmen", "last": "Dalton", "title": "Manager, Engineering & Maintenance Category", "company": "americold.com", "phone": "(706) 693-4100", "email": "carmen.dalton@americold.com", "city": "Gainesville", "state": "Florida", "linkedin": "https://www.linkedin.com/in/carmen-dalton-279a9161"}, {"id": "1581538871", "first": "William", "last": "Kane", "title": "Manager, Maintenance & Reliability", "company": "gp.com", "phone": "", "email": "william.kane@gapac.com", "city": "", "state": "Florida", "linkedin": "https://www.linkedin.com/in/william-kane-a32175155"}, {"id": "5253769797", "first": "Jonathon", "last": "Causey", "title": "Manager, Maintenance", "company": "mohawkind.com", "phone": "(336) 313-4022 ext. 54022", "email": "john_causey@mohawkind.com", "city": "Greensboro", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/john-causey-3520b67"}, {"id": "7873566275", "first": "Cameron", "last": "Sloane", "title": "Maintenance Manager, Operations", "company": "walmart.com", "phone": "(334) 705-2850", "email": "cameron.sloane@walmart.com", "city": "Auburn", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/cameron-sloane-b6b51211b"}, {"id": "3416315608", "first": "Tony", "last": "Threadgill", "title": "Process Manager Maintenance", "company": "internationalpaper.com", "phone": "(334) 418-5300", "email": "tony.threadgill@ipaper.com", "city": "Selma", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/tony-threadgill-449284138"}, {"id": "9452575597", "first": "Hugh", "last": "Cox", "title": "Manager, Facility Maintenance", "company": "onelineage.com", "phone": "(256) 744-6917", "email": "hcox@onelineage.com", "city": "Decatur", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/hugh-cox-134705318"}, {"id": "-1987244097", "first": "Bill", "last": "Wilson", "title": "Manager, Maintenance", "company": "teknorapex.com", "phone": "(239) 337-0400", "email": "bill.wilson@teknorapex.com", "city": "Fort Myers", "state": "Florida", "linkedin": "https://www.linkedin.com/in/bill-wilson-3523b684"}, {"id": "-2091692465", "first": "Olivia", "last": "Spivey", "title": "Channel Facilities Care Solutions Key Account Manager", "company": "3m.com", "phone": "(561) 449-1185", "email": "ospivey@3m.com", "city": "Boca Raton", "state": "Florida", "linkedin": "https://www.linkedin.com/in/olivia-spivey-9aaa9447"}, {"id": "-1918394014", "first": "Christopher", "last": "Baxley", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "", "email": "christopher.baxley@walmart.com", "city": "Hope Mills", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/chris-baxley-09752436"}, {"id": "-1320106937", "first": "Clay", "last": "Mobley", "title": "Manager, Maintenance", "company": "kindermorgan.com", "phone": "(910) 813-5283", "email": "clay_mobley@kindermorgan.com", "city": "Chester", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/clay-mobley-2b1362106"}, {"id": "-2137555347", "first": "Charity", "last": "Carlson", "title": "Project Controls Office of Infrastructure & Facilities Management Engineer", "company": "honeywell.com", "phone": "(816) 997-5248", "email": "charity.carlson@honeywell.com", "city": "Charlotte", "state": "North Carolina", "linkedin": ""}, {"id": "7103215581", "first": "Dawson", "last": "Stacy", "title": "Manager, Maintenance", "company": "pepsico.com", "phone": "(727) 608-7803", "email": "dawson.stacy@pepsico.com", "city": "Tampa", "state": "Florida", "linkedin": "https://www.linkedin.com/in/dawson-stacy-1026bb141"}, {"id": "2634384747", "first": "Christopher", "last": "Volkers", "title": "Manager, Maintenance", "company": "mauserpackaging.com", "phone": "(216) 361-8958", "email": "chrisv@mauserpackaging.com", "city": "Mount Dora", "state": "Florida", "linkedin": "https://www.linkedin.com/in/chris-volkers-b43778b6"}, {"id": "7037642286", "first": "Jay", "last": "Hucks", "title": "Manager, Maintenance", "company": "domtar.com", "phone": "(843) 523-5058", "email": "jay.hucks@domtar.com", "city": "Fort Mill", "state": "South Carolina", "linkedin": ""}, {"id": "5424170382", "first": "Daniel", "last": "Howard", "title": "Area Manager Rolling Mill Maintenance", "company": "cmc.com", "phone": "(803) 936-3761", "email": "daniel.howard@cmc.com", "city": "Columbia", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/daniel-howard-939291133"}, {"id": "5646227381", "first": "Shawn", "last": "Kingston", "title": "Manager, Maintenance", "company": "teknorapex.com", "phone": "(731) 519-1329", "email": "skingston@teknorapex.com", "city": "Brownsville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/sean-kingston-4696b0255"}, {"id": "2410608577", "first": "Brody", "last": "Huber", "title": "Manager, Maintenance & Engineering", "company": "nouryon.com", "phone": "(423) 629-1405 ext. 3344", "email": "brody.huber@nouryon.com", "city": "Axis", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/brody-huber-a11294160"}, {"id": "3268217158", "first": "Allen", "last": "Swann", "title": "Maintenance Manager, Operations", "company": "walmart.com", "phone": "(256) 737-3190", "email": "allen.swann@walmart.com", "city": "Cullman", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/allen-swann-0ba964203"}, {"id": "5116630474", "first": "Andrew", "last": "Lorenz", "title": "Senior Project Engineer & Manager, Construction", "company": "bridgestoneamericas.com", "phone": "(615) 477-2791", "email": "lorenzandrew@bfusa.com", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/andrew-lorenz-033b2341"}, {"id": "1929829197", "first": "Robert", "last": "Hallman", "title": "Area Maintenance Manager", "company": "internationalpaper.com", "phone": "(803) 353-7083", "email": "robert.hallman@ipaper.com", "city": "Springfield", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/robert-hallman-813a68117"}, {"id": "14676503379", "first": "Victor", "last": "Otero", "title": "Senior Manager, Facilities", "company": "walmart.com", "phone": "(407) 826-6989", "email": "victor.otero@walmart.com", "city": "DeBary", "state": "Florida", "linkedin": "https://www.linkedin.com/in/victor-otero-76b08117"}, {"id": "1882583267", "first": "Lynn", "last": "Paczkowski", "title": "Manager, Information Technology Source & Plant Maintenance", "company": "internationalpaper.com", "phone": "(901) 419-8738", "email": "lynn.paczkowski@ipaper.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/lynn-paczkowski-a5797023"}, {"id": "-2059010356", "first": "Lyons", "last": "Aaron", "title": "Manager, Maintenance", "company": "republicservices.com", "phone": "(702) 465-1547", "email": "lyo@republicservices.com", "city": "Union City", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/lyons-aaron-3063a8a1"}, {"id": "-1006763718", "first": "Marcus", "last": "Houston", "title": "Facility Assistant & Manager", "company": "ergon.com", "phone": "(901) 774-0068", "email": "marcus.houston@ergon.com", "city": "Memphis", "state": "Tennessee", "linkedin": ""}, {"id": "9638956231", "first": "Courtney", "last": "Presnell", "title": "Account Manager, Facility Coding", "company": "3m.com", "phone": "(262) 844-6749", "email": "cpresnell@mmm.com", "city": "Clyde", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/courtney-presnell-1533b615"}, {"id": "1853252941", "first": "Kevin", "last": "Chiplock", "title": "Operations Manager, Maintenance", "company": "walmart.com", "phone": "(386) 418-5968", "email": "kevin.chiplock@walmart.com", "city": "High Springs", "state": "Florida", "linkedin": "https://www.linkedin.com/in/kevin-chiplock-3b70a616"}, {"id": "3389687609", "first": "Braxton", "last": "Swanner", "title": "Manager, Facilities", "company": "walmart.com", "phone": "(256) 282-7881", "email": "braxton.swanner@walmart.com", "city": "Huntsville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/braxton-swanner-09306517b"}, {"id": "16113848430", "first": "Chase", "last": "Garner", "title": "Manager, Facility Environmental", "company": "weyerhaeuser.com", "phone": "(205) 391-8757", "email": "chase.garner@weyerhaeuser.com", "city": "Tuscaloosa", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/chase-garner-a3006b17b"}, {"id": "-2010536118", "first": "Perry", "last": "Tyler", "title": "Project Manager, Capital", "company": "internationalpaper.com", "phone": "(757) 569-4138", "email": "perry.tyler@ipaper.com", "city": "Gates", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/perry-tyler-84597b90"}, {"id": "1755848849", "first": "Michael", "last": "Marion", "title": "Maintenance Reliability Manager", "company": "internationalpaper.com", "phone": "(334) 361-5058", "email": "mike.marion@ipaper.com", "city": "Deatsville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/mike-marion-b7624635"}, {"id": "1444786332", "first": "John", "last": "Sylvina", "title": "Senior Manager, Retail Facilities", "company": "bridgestoneamericas.com", "phone": "(610) 968-2038", "email": "jsylvina@bfrc.com", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/john-sylvina-2b37388"}, {"id": "2359502268", "first": "Jeffrey", "last": "Pittman", "title": "Assistant Maintenance Manager", "company": "internationalpaper.com", "phone": "(910) 887-6263", "email": "jeffrey.pittman@ipaper.com", "city": "Lumberton", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/jeffrey-pittman-6008a392"}, {"id": "2077202722", "first": "Alan", "last": "Byers", "title": "Manager, Facility & Maintenance", "company": "honeywell.com", "phone": "(843) 301-6646", "email": "abyers@honeywell.com", "city": "Isle of Palms", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/alan-byers-78853013"}, {"id": "9693947398", "first": "Steven", "last": "Hopkins", "title": "Manager I, Facility Operations", "company": "safety-kleen.com", "phone": "(803) 520-2866", "email": "steve.hopkins@safety-kleen.com", "city": "Lexington", "state": "South Carolina", "linkedin": ""}, {"id": "6908095485", "first": "Wyatt", "last": "Robertson", "title": "Manager, Maintenance", "company": "walmart.com", "phone": "(336) 552-3423", "email": "wyatt.robertson@walmart.com", "city": "Eden", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/wyatt-robertson-2377a4158"}, {"id": "-2049375007", "first": "Steve", "last": "Latta", "title": "Manager, Maintenance", "company": "mauserpackaging.com", "phone": "(256) 831-8441", "email": "steve.latta@mauserpackaging.com", "city": "Anniston", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/steve-latta-53865962"}, {"id": "5166653358", "first": "Henry", "last": "Clement", "title": "Manager, Network Maintenance", "company": "corporate.comcast.com", "phone": "(786) 566-6940", "email": "heclement@corporate.comcast.com", "city": "Miami", "state": "Florida", "linkedin": "https://www.linkedin.com/in/henry-clement-55044b98"}, {"id": "1900850917", "first": "Joseph", "last": "Burns", "title": "Facility Service Manager", "company": "walmart.com", "phone": "(609) 306-8813", "email": "joe.burns@walmart.com", "city": "Leland", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/joseph-sc-burns"}, {"id": "2085937191", "first": "Brian", "last": "Schaffer", "title": "Manager, Maintenance", "company": "kraton.com", "phone": "(731) 499-0863", "email": "brian.schaffer@kraton.com", "city": "Lynn Haven", "state": "Florida", "linkedin": "https://www.linkedin.com/in/brianschaffer11"}, {"id": "12322424488", "first": "Jonathan", "last": "Bryant", "title": "Manager, Maintenance", "company": "kochfoods.com", "phone": "(423) 328-6904", "email": "john.bryant@kochfoods.com", "city": "Afton", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/john-bryant-56601870"}, {"id": "12977201133", "first": "Michael", "last": "Cobb", "title": "Operation Confined Space High Angle Hazmat Fire Team Manager, Facility & Instructor & Project Manager, Remediation & Hhw", "company": "cleanharbors.com", "phone": "(870) 864-2296", "email": "cobb.michael@cleanharbors.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/michael-cobb-11bb64198"}, {"id": "7453080876", "first": "Clinton", "last": "Walker", "title": "Manager, Maintenance Engineering", "company": "kraftheinzcompany.com", "phone": "(410) 979-9289", "email": "clint.walker@kraftheinz.com", "city": "Fort Myers", "state": "Florida", "linkedin": "https://www.linkedin.com/in/clinton-walker-0714b124"}, {"id": "15575262719", "first": "Shawn", "last": "Ebling", "title": "Manager, Maintenance", "company": "landolakesinc.com", "phone": "(816) 558-3527", "email": "sdebling@landolakes.com", "city": "Nashville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/shawn-ebling-3632a470"}, {"id": "2465598407", "first": "David", "last": "Dietel", "title": "Manager, Operations & Maintenance", "company": "rwe.com", "phone": "(737) 239-5241", "email": "david.dietel@rwe.com", "city": "", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/david-dietel-98aa02103"}, {"id": "2424705544", "first": "Jeffrey", "last": "Frye", "title": "Fibers E & I Maintenance Manager", "company": "internationalpaper.com", "phone": "(803) 353-7507", "email": "jeff.frye@ipaper.com", "city": "Lexington", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/jeff-frye-a10306b8"}, {"id": "2517647614", "first": "Jeffrey", "last": "Wallace", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(407) 888-1659", "email": "jeff.wallace@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": "https://www.linkedin.com/in/jeff-wallace-3319781a2"}, {"id": "2050716497", "first": "Connie", "last": "Dickerson", "title": "Manager, Facilities", "company": "walmart.com", "phone": "(901) 743-6210", "email": "connie.dickerson@walmart.com", "city": "", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/connie-dickerson-b9783550"}, {"id": "10859592538", "first": "Roger", "last": "Witt", "title": "Manager, Maintenance", "company": "kochfoods.com", "phone": "(256) 549-6239", "email": "roger.witt@kochfoods.com", "city": "Gadsden", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/roger-witt-09358955"}, {"id": "1849738501", "first": "David", "last": "Sobol", "title": "Manager, Maintenance Eh & S", "company": "ball.com", "phone": "(252) 984-4100", "email": "david.sobol@ball.com", "city": "Reidsville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/david-sobol-58963557"}, {"id": "5613837064", "first": "Jason", "last": "Ward", "title": "Manager, Area Maintenace & Facilities", "company": "7-eleven.com", "phone": "(252) 382-0643", "email": "jason.ward@7-11.com", "city": "Wilson", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/jason-ward87"}, {"id": "2657564107", "first": "David", "last": "Klimek", "title": "Project Manager, Capital Improvement", "company": "internationalpaper.com", "phone": "(334) 418-5164", "email": "david.klimek@ipaper.com", "city": "Prattville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/david-klimek-520884a8"}, {"id": "9678225027", "first": "Shaun", "last": "Reynolds", "title": "Manager, Engineering & Maintenance", "company": "americold.com", "phone": "(207) 450-4315", "email": "shaun.reynolds@americold.com", "city": "Kingston", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/shaun-reynolds-072b9617b"}, {"id": "3356886004", "first": "Matthew", "last": "Ziglar", "title": "Area Maintenance Manager Fibers", "company": "internationalpaper.com", "phone": "(850) 968-3002", "email": "matthew.ziglar@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": ""}, {"id": "412013017", "first": "Thomas", "last": "Massey", "title": "Manager, Maintenance", "company": "nucor.com", "phone": "(903) 626-6217", "email": "tommy.massey@nucor.com", "city": "Middleburg", "state": "Florida", "linkedin": "https://www.linkedin.com/in/thomas-massey-643356b2"}, {"id": "7102697381", "first": "Acker", "last": "Mike", "title": "Manager, Maintenance", "company": "internationalpaper.com", "phone": "(651) 645-0811", "email": "acker.mike@internationalpaper.com", "city": "Carrollton", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/acker-mike-183a289a"}, {"id": "3341878228", "first": "Stephen", "last": "Bibb", "title": "Area Maintenance Manager", "company": "internationalpaper.com", "phone": "(334) 361-5097", "email": "steve.bibb@ipaper.com", "city": "Prattville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/steve-bibb-496815225"}, {"id": "-2033630978", "first": "Chad", "last": "Kaunitz", "title": "E Maintenance Technology & Valve Asset Manager I", "company": "internationalpaper.com", "phone": "(850) 968-2121 ext. 2527", "email": "chad.kaunitz@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": "https://www.linkedin.com/in/chad-kaunitz-883a017a"}, {"id": "3051676855", "first": "Ali", "last": "Elribie", "title": "Manager, Facilities Operations", "company": "aramark.com", "phone": "(704) 576-5391", "email": "elribie-ali@aramark.com", "city": "Charlotte", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/ali-elribie-b5a2654b"}, {"id": "9865763874", "first": "Adam", "last": "Colwell", "title": "Manager, Operations Maintenance", "company": "walmart.com", "phone": "(541) 401-3249", "email": "adam.colwell@walmart.com", "city": "Winter Haven", "state": "Florida", "linkedin": "https://www.linkedin.com/in/adam-colwell-b4aa24195"}, {"id": "2548535676", "first": "Marvin", "last": "Gainey", "title": "Manager, Maintenance", "company": "mohawkind.com", "phone": "(843) 309-9278", "email": "marvin-g@mohawkind.com", "city": "Hartsville", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/marvin-gainey-0b791039"}, {"id": "2362278044", "first": "Brandon", "last": "Muncy", "title": "Manager, Maintenance", "company": "conagrabrands.com", "phone": "(479) 498-7539", "email": "brandon.muncy@conagrabrands.com", "city": "Russellville", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/brandon-muncy-0b61abb7"}, {"id": "7080815540", "first": "Zacariah", "last": "Wilson", "title": "Project Manager, Construction", "company": "honda.com", "phone": "(256) 294-0351", "email": "zacariah_wilson@ahm.honda.com", "city": "Anniston", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/zacariah-wilson-80662520a"}, {"id": "5949839726", "first": "Eligio", "last": "Muñiz", "title": "Senior Manager, Facilities", "company": "walmart.com", "phone": "(479) 935-0712", "email": "eligio.muniz@walmart.com", "city": "Orlando", "state": "Florida", "linkedin": "https://www.linkedin.com/in/eligio-mu%c3%b1iz"}, {"id": "2366113932", "first": "Jimmie", "last": "Lawson", "title": "Maintenance Manager, Operations", "company": "walmart.com", "phone": "(704) 487-2572", "email": "jimmie.lawson@walmart.com", "city": "Spartanburg", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/jimmie-lawson-8376888a"}, {"id": "-917494279", "first": "Clayton", "last": "Rachels", "title": "Manager, Program Maintenance", "company": "norfolksouthern.com", "phone": "(731) 514-1518", "email": "clayton.rachels@nscorp.com", "city": "Trenton", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/clay-rachels-16b2693a"}, {"id": "3755014534", "first": "Tyler", "last": "Waters", "title": "Manager, Maintenance & Reliability", "company": "mohawkind.com", "phone": "", "email": "tyler_waters@mohawkind.com", "city": "Fuquay-Varina", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/tyler-waters-a958683"}, {"id": "10817001788", "first": "Benjamin", "last": "Jacoby", "title": "Senior Manager, Facilities", "company": "domtar.com", "phone": "(704) 957-9016", "email": "benjamin.jacoby@domtar.com", "city": "Fort Mill", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/benjamin-jacoby-aoee-leed-green-associate-b3909259"}, {"id": "1953987270", "first": "John", "last": "Cornelison", "title": "Senior Manager, Regional Maintenance", "company": "onelineage.com", "phone": "(256) 538-2653", "email": "jcornelison@lineagelogistics.com", "city": "Attalla", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/john-cornelison-7a95b122"}, {"id": "2088000552", "first": "Vincent", "last": "Guidi-Louviere", "title": "Senior Project Manager, Major Capital Projects", "company": "domtar.com", "phone": "(434) 534-1990", "email": "vincent.guidi@domtar.com", "city": "Greer", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/vincent-guidi-louviere-pmp-7b040535"}, {"id": "8113144142", "first": "Angela", "last": "Shearin", "title": "Operation Manager, Maintenance & Assistant", "company": "walmart.com", "phone": "(252) 430-5651", "email": "angela.shearin@walmart.com", "city": "Henderson", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/angela-shearin-60a568141"}, {"id": "2661355640", "first": "Clark", "last": "Johnson", "title": "Area Maintenance Manager Pulp", "company": "internationalpaper.com", "phone": "(334) 418-5446", "email": "clark.johnson@ipaper.com", "city": "Montgomery", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/clark-johnson-79166b38"}, {"id": "2005154021", "first": "Bruce", "last": "Gramling", "title": "Manager, Maintenance & Reliability", "company": "gp.com", "phone": "(803) 718-0904", "email": "bruce.gramling@gapac.com", "city": "Fort Mill", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/brucegramling"}, {"id": "7463706534", "first": "Christopher", "last": "Willson", "title": "Manager, Engineering & Maintenance", "company": "americold.com", "phone": "(309) 457-1419", "email": "christopher.willson@americold.com", "city": "Fayetteville", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/christopher-willson-b770121ab"}, {"id": "2964284919", "first": "Paul", "last": "Turner", "title": "Refrigeration and Maintenance Manager", "company": "kochfoods.com", "phone": "(256) 549-6599 ext. 2535", "email": "paul.turner@kochfoods.com", "city": "Collinsville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/paul-turner-855284114"}, {"id": "5954167896", "first": "Michael", "last": "Drewry", "title": "Process Manager Maintenance", "company": "internationalpaper.com", "phone": "(334) 418-5506", "email": "michael.drewry@ipaper.com", "city": "Maplesville", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/michael-drewry-7919aa162"}, {"id": "-2034363583", "first": "Allen", "last": "Tidwell", "title": "Manager, Facilities Services", "company": "americold.com", "phone": "(803) 799-6520 ext. 4", "email": "allen.tidwell@americold.com", "city": "", "state": "South Carolina", "linkedin": "https://www.linkedin.com/in/allen-tidwell-60b75467"}, {"id": "-2036899370", "first": "Lawrence", "last": "Coon", "title": "Area Maintenance Manager Power", "company": "internationalpaper.com", "phone": "(850) 968-3007", "email": "lawrence.coon@ipaper.com", "city": "Orange Park", "state": "Florida", "linkedin": "https://www.linkedin.com/in/lawrence-coon-8617b13a"}, {"id": "2789910954", "first": "Larry", "last": "Triplett", "title": "Manager, Maintenance", "company": "ppg.com", "phone": "(256) 859-2500", "email": "larrytriplett@ppg.com", "city": "", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/larry-triplett-8b72a374"}, {"id": "57867334", "first": "Donny", "last": "Stewart", "title": "Manager, Facility", "company": "internationalpaper.com", "phone": "(901) 748-5107", "email": "donny.stewart@ipaper.com", "city": "Memphis", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/donny-stewart-a4750410"}, {"id": "-2059095138", "first": "Richard", "last": "Dunaway", "title": "Fleet Maintenance Manager", "company": "republicservices.com", "phone": "(423) 867-6560", "email": "rdunaway@republicservices.com", "city": "Ooltewah", "state": "Tennessee", "linkedin": "https://www.linkedin.com/in/richard-dunaway-452184109"}, {"id": "2011133989", "first": "Nelson", "last": "Roman", "title": "Senior Manager, Facilities", "company": "jnj.com", "phone": "(904) 443-1000", "email": "nroman3@its.jnj.com", "city": "Miramar", "state": "Florida", "linkedin": "https://www.linkedin.com/in/nelson-roman-0006aa277"}, {"id": "1707171712", "first": "John", "last": "Thompson", "title": "Manager, Maintenance", "company": "weyerhaeuser.com", "phone": "(205) 596-1148", "email": "jthompson@weyerhaeuser.com", "city": "Millport", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/john-thompson-b6951a40"}, {"id": "-1254020118", "first": "Doug", "last": "Livingston", "title": "Senior Facility Service Manager", "company": "americold.com", "phone": "(813) 754-9341 ext. 54068", "email": "livingston@americold.com", "city": "Fort Lauderdale", "state": "Florida", "linkedin": "https://www.linkedin.com/in/doug-livingston-20767918"}, {"id": "3264148669", "first": "Timothy", "last": "Chandler", "title": "Manager, Facility Capital, Riverdale Mill", "company": "internationalpaper.com", "phone": "(912) 238-7850", "email": "timothy.chandler@ipaper.com", "city": "Montgomery", "state": "Alabama", "linkedin": "https://www.linkedin.com/in/tchand"}, {"id": "9677784416", "first": "Travis", "last": "Greene", "title": "Manager, Maintenance", "company": "nucor.com", "phone": "(903) 626-6217", "email": "greene@vulcraft.com", "city": "Clemmons", "state": "North Carolina", "linkedin": "https://www.linkedin.com/in/travis-greene-5ba1847b"}] ;

const STATUSES = ["not called","called - no answer","left voicemail","wrong number","sent email","spoke - interested","spoke - not interested","meeting set","follow up","deal closed","dead"];
const STATUS_COLORS = {
  "not called": "#888780",
  "called - no answer": "#B4B2A9",
  "left voicemail": "#AFA9EC",
  "wrong number": "#C47A3A",
  "sent email": "#378ADD",
  "spoke - interested": "#1D9E75",
  "spoke - not interested": "#E24B4A",
  "meeting set": "#9B59B6",
  "follow up": "#EF9F27",
  "deal closed": "#0F6E56",
  "dead": "#444441",
};
const STORAGE_KEY = "crm_data_v1";
const CONTACTS_KEY = "crm_contacts_v1";
const SALESPEOPLE_KEY = "crm_salespeople_v1";
const ACTIVE_SP_KEY = "crm_active_sp_v1";

// Legacy keys kept only for one-time migration of sp1's data
const LEGACY_ADDED_KEY = "crm_added_contacts_v1";
const LEGACY_DELETED_KEY = "crm_deleted_ids_v1";

function defaultSalespeople() {
  return Array.from({ length: 15 }, (_, i) => ({ id: `sp${i + 1}`, name: `Salesperson ${i + 1}` }));
}

function initContacts(spId) {
  const key = `${CONTACTS_KEY}_${spId}`;
  try {
    const saved = localStorage.getItem(key);
    if (saved) return JSON.parse(saved);
  } catch {}
  // One-time migration for sp1: reconstruct from RAW_CONTACTS + legacy added/deleted
  if (spId === 'sp1') {
    try {
      const added = JSON.parse(localStorage.getItem(`${LEGACY_ADDED_KEY}_sp1`) || localStorage.getItem(LEGACY_ADDED_KEY) || '[]');
      const deleted = new Set(JSON.parse(localStorage.getItem(`${LEGACY_DELETED_KEY}_sp1`) || localStorage.getItem(LEGACY_DELETED_KEY) || '[]'));
      const migrated = [...RAW_CONTACTS.filter(c => !deleted.has(c.id)), ...added];
      localStorage.setItem(key, JSON.stringify(migrated));
      return migrated;
    } catch {}
    return [...RAW_CONTACTS];
  }
  return [];
}

function initData(key) {
  try {
    const saved = localStorage.getItem(key);
    if (saved) return JSON.parse(saved);
    // migrate legacy data for the first salesperson
    if (key !== STORAGE_KEY) {
      const legacy = localStorage.getItem(STORAGE_KEY);
      if (legacy) return JSON.parse(legacy);
    }
  } catch (e) {
    return {};
  }
  const map = {};
  RAW_CONTACTS.forEach((c) => {
    map[c.id] = { status: "not called", notes: "", followUpDate: "", lastCalled: "", calls: [] };
  });
  return map;
}

function saveData(data, key) {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    // ignore
  }
}

function initials(first, last) {
  const a = (first || "")[0] || (last || "")[0] || "?";
  const b = (last || "")[0] || "";
  return `${a.toUpperCase()}${b.toUpperCase()}`;
}

const COMPANY_COLORS = ["#7F77DD", "#1D9E75", "#D85A30", "#D4537E", "#378ADD", "#639922", "#BA7517", "#E24B4A"];
const companyColorMap = {};
let colorIdx = 0;
function companyColor(co) {
  if (!companyColorMap[co]) {
    companyColorMap[co] = COMPANY_COLORS[colorIdx++ % COMPANY_COLORS.length];
  }
  return companyColorMap[co];
}

const outerGridStyle = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(120px,1fr))',
  gap: '10px',
  marginBottom: '1.5rem',
};

const cardBaseStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
  padding: '10px 14px',
  border: '0.5px solid var(--color-border-tertiary)',
  borderRadius: 'var(--border-radius-md)',
  cursor: 'pointer',
  background: 'var(--color-background-primary)',
  transition: 'border-color 0.1s',
};

export default function CRM() {
  const [salespeople, setSalespeople] = useState(() => {
    try {
      const saved = localStorage.getItem(SALESPEOPLE_KEY);
      if (saved) return JSON.parse(saved);
    } catch {}
    const defaults = defaultSalespeople();
    localStorage.setItem(SALESPEOPLE_KEY, JSON.stringify(defaults));
    return defaults;
  });
  const [activeSP, setActiveSP] = useState(() => localStorage.getItem(ACTIVE_SP_KEY) || 'sp1');
  const [showManageTeam, setShowManageTeam] = useState(false);
  const [editingNames, setEditingNames] = useState([]);

  const [crmData, setCrmData] = useState(() => initData(`${STORAGE_KEY}_${localStorage.getItem(ACTIVE_SP_KEY) || 'sp1'}`));
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterCompany, setFilterCompany] = useState("all");
  const [sortBy, setSortBy] = useState("name");
  const [selected, setSelected] = useState(null);
  const [editNotes, setEditNotes] = useState("");
  const [editFollowUp, setEditFollowUp] = useState("");
  const [editStatus, setEditStatus] = useState("not called");
  const [callNote, setCallNote] = useState("");
  const [view, setView] = useState("list");
  const [tab, setTab] = useState("overview");
  const [contacts, setContacts] = useState(() => initContacts(localStorage.getItem(ACTIVE_SP_KEY) || 'sp1'));
  const [importResult, setImportResult] = useState("");
  const fileInputRef = useRef(null);

  const spDataKey = `${STORAGE_KEY}_${activeSP}`;
  const spContactsKey = `${CONTACTS_KEY}_${activeSP}`;

  useEffect(() => {
    localStorage.setItem(ACTIVE_SP_KEY, activeSP);
    setCrmData(initData(spDataKey));
    setContacts(initContacts(activeSP));
    setSelected(null);
  }, [activeSP]); // eslint-disable-line react-hooks/exhaustive-deps

  const allContacts = contacts;

  const companies = useMemo(() => {
    const set = new Set(allContacts.map((c) => c.company).filter(Boolean));
    return Array.from(set).sort();
  }, [allContacts]);

  const filtered = useMemo(() => {
    let list = allContacts;
    if (search) {
      const q = search.toLowerCase();
      list = list.filter((c) => {
        const phone = (c.phone || "").toLowerCase();
        return (`${c.first} ${c.last}`.toLowerCase().includes(q) ||
          c.company.toLowerCase().includes(q) ||
          c.title.toLowerCase().includes(q) ||
          phone.includes(q) ||
          c.email.toLowerCase().includes(q));
      });
    }
    if (filterStatus !== "all") {
      list = list.filter((c) => (crmData[c.id]?.status || "not called") === filterStatus);
    }
    if (filterCompany !== "all") {
      list = list.filter((c) => c.company === filterCompany);
    }
    if (sortBy === "name") {
      list = [...list].sort((a, b) => a.last.localeCompare(b.last));
    } else if (sortBy === "company") {
      list = [...list].sort((a, b) => a.company.localeCompare(b.company));
    } else if (sortBy === "followup") {
      list = [...list].sort((a, b) => {
        const da = crmData[a.id]?.followUpDate || "9999";
        const db = crmData[b.id]?.followUpDate || "9999";
        return da.localeCompare(db);
      });
    } else if (sortBy === "status") {
      list = [...list].sort((a, b) => {
        const sa = STATUSES.indexOf(crmData[a.id]?.status || "not called");
        const sb = STATUSES.indexOf(crmData[b.id]?.status || "not called");
        return sa - sb;
      });
    }
    return list;
  }, [search, filterStatus, filterCompany, sortBy, crmData, allContacts]);

  const stats = useMemo(() => {
    const counts = {};
    STATUSES.forEach((s) => { counts[s] = 0; });
    allContacts.forEach((c) => {
      const s = crmData[c.id]?.status || "not called";
      counts[s] = (counts[s] || 0) + 1;
    });
    return counts;
  }, [crmData, allContacts]);

  const todayFollowUps = useMemo(() => {
    const today = new Date().toISOString().split('T')[0];
    return allContacts.filter((c) => crmData[c.id]?.followUpDate === today);
  }, [crmData, allContacts]);

  const upcomingFollowUps = useMemo(() => {
    const today = new Date().toISOString().split('T')[0];
    return allContacts.filter((c) => {
      const d = crmData[c.id]?.followUpDate;
      return d && d >= today;
    }).sort((a, b) => (crmData[a.id]?.followUpDate || "").localeCompare(crmData[b.id]?.followUpDate || ""));
  }, [crmData, allContacts]);

  function openContact(contact) {
    setSelected(contact);
    const data = crmData[contact.id] || {};
    setEditNotes(data.notes || "");
    setEditFollowUp(data.followUpDate || "");
    setEditStatus(data.status || "not called");
    setCallNote("");
    setTab("overview");
  }

  function saveContact() {
    if (!selected) return;
    const updated = {
      ...crmData,
      [selected.id]: {
        ...crmData[selected.id],
        notes: editNotes,
        followUpDate: editFollowUp,
        status: editStatus,
      },
    };
    setCrmData(updated);
    saveData(updated, spDataKey);
  }

  function logCall() {
    if (!selected) return;
    if (!callNote.trim() && editStatus === (crmData[selected.id]?.status || "not called")) return;
    const now = new Date();
    const dateStr = now.toLocaleDateString('en-US');
    const timeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const existing = crmData[selected.id] || {};
    const calls = [...(existing.calls || []), { date: dateStr, time: timeStr, note: callNote, status: editStatus }];
    const updated = {
      ...crmData,
      [selected.id]: {
        ...existing,
        notes: editNotes,
        followUpDate: editFollowUp,
        status: editStatus,
        lastCalled: dateStr,
        calls,
      },
    };
    setCrmData(updated);
    saveData(updated, spDataKey);
    setCallNote("");
  }

  function logQuickCall(note, status) {
    if (!selected) return;
    const now = new Date();
    const dateStr = now.toLocaleDateString('en-US');
    const timeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const existing = crmData[selected.id] || {};
    const calls = [...(existing.calls || []), { date: dateStr, time: timeStr, note, status }];
    const updated = {
      ...crmData,
      [selected.id]: {
        ...existing,
        notes: editNotes,
        followUpDate: editFollowUp,
        status,
        lastCalled: dateStr,
        calls,
      },
    };
    setEditStatus(status);
    setCrmData(updated);
    saveData(updated, spDataKey);
  }

  function deleteContact(id) {
    if (!window.confirm('Remove this contact? This cannot be undone.')) return;
    const updated = contacts.filter(c => c.id !== id);
    setContacts(updated);
    localStorage.setItem(spContactsKey, JSON.stringify(updated));
    setSelected(null);
  }

  function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const norm = (s) => String(s || '').toLowerCase().replace(/[\s_\-]+/g, '');
        const pick = (row, ...keys) => {
          for (const k of keys) {
            for (const rk of Object.keys(row)) {
              if (norm(rk) === norm(k)) return String(row[rk] || '').trim();
            }
          }
          return '';
        };
        const newContacts = rows.map((row, i) => ({
          id: `imported_${Date.now()}_${i}`,
          first: pick(row, 'first', 'firstname', 'first name'),
          last: pick(row, 'last', 'lastname', 'last name'),
          title: pick(row, 'title', 'job title', 'jobtitle', 'position'),
          company: pick(row, 'company', 'organization', 'employer'),
          phone: pick(row, 'phone', 'phone number', 'phonenumber', 'mobile'),
          email: pick(row, 'email', 'email address', 'emailaddress'),
          city: pick(row, 'city'),
          state: pick(row, 'state', 'province'),
          linkedin: pick(row, 'linkedin', 'linkedin url', 'linkedinurl'),
        })).filter(c => c.first || c.last);
        const updated = [...contacts, ...newContacts];
        setContacts(updated);
        localStorage.setItem(spContactsKey, JSON.stringify(updated));
        setImportResult(`✓ Imported ${newContacts.length} contact${newContacts.length !== 1 ? 's' : ''}`);
        setTimeout(() => setImportResult(""), 4000);
      } catch {
        setImportResult('✗ Could not parse file — check the format');
        setTimeout(() => setImportResult(""), 4000);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  }

  const callStats = useMemo(() => {
    const now = new Date();
    const todayStr = now.toLocaleDateString('en-US');
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - now.getDay());
    startOfWeek.setHours(0, 0, 0, 0);
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    let todayCount = 0, weekCount = 0, monthCount = 0;
    const todayCalls = [];
    const monthByOutcome = {};
    Object.entries(crmData).forEach(([id, entry]) => {
      (entry.calls || []).forEach(call => {
        const callDate = new Date(call.date);
        if (call.date === todayStr) {
          todayCount++;
          const c = allContacts.find(x => x.id === id);
          if (c) todayCalls.push({ name: `${c.first} ${c.last}`, time: call.time, note: call.note, status: call.status });
        }
        if (callDate >= startOfWeek) weekCount++;
        if (callDate >= startOfMonth) {
          monthCount++;
          monthByOutcome[call.status] = (monthByOutcome[call.status] || 0) + 1;
        }
      });
    });
    todayCalls.sort((a, b) => b.time.localeCompare(a.time));
    return { todayCount, weekCount, monthCount, todayCalls, monthByOutcome };
  }, [crmData, allContacts]);

  const battleStats = useMemo(() => {
    const now = new Date();
    const todayStr = now.toLocaleDateString('en-US');
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - now.getDay());
    startOfWeek.setHours(0, 0, 0, 0);
    const startOfLastWeek = new Date(startOfWeek);
    startOfLastWeek.setDate(startOfLastWeek.getDate() - 7);
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    return salespeople.map(sp => {
      let spData = sp.id === activeSP ? crmData : {};
      if (sp.id !== activeSP) {
        try {
          const saved = localStorage.getItem(`${STORAGE_KEY}_${sp.id}`);
          if (saved) spData = JSON.parse(saved);
        } catch {}
      }
      let today = 0, thisWeek = 0, lastWeek = 0, thisMonth = 0;
      Object.values(spData).forEach(entry => {
        (entry.calls || []).forEach(call => {
          const d = new Date(call.date);
          if (call.date === todayStr) today++;
          if (d >= startOfWeek) thisWeek++;
          if (d >= startOfLastWeek && d < startOfWeek) lastWeek++;
          if (d >= startOfMonth) thisMonth++;
        });
      });
      return { id: sp.id, name: sp.name, today, thisWeek, lastWeek, thisMonth };
    }).sort((a, b) => b.today - a.today || b.thisWeek - a.thisWeek);
  }, [salespeople, crmData, activeSP]);

  const isToday = (date) => date === new Date().toISOString().split('T')[0];
  const isPast = (date) => date && date < new Date().toISOString().split('T')[0];

  return (
    <>
    <div style={{ padding: '1rem 0', fontFamily: 'var(--font-sans)' }}>
      <h2 style={{ position: 'absolute', left: '-9999px', top: 'auto', width: '1px', height: '1px', overflow: 'hidden' }}>Sales CRM — {allContacts.length} contacts</h2>

      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap', marginBottom: '1rem', paddingBottom: '10px', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
        <span style={{ fontSize: '12px', color: 'var(--color-text-secondary)', marginRight: '2px', flexShrink: 0 }}>viewing as:</span>
        {salespeople.map(sp => (
          <button
            key={sp.id}
            onClick={() => setActiveSP(sp.id)}
            style={{
              padding: '4px 12px',
              borderRadius: '20px',
              border: `0.5px solid ${activeSP === sp.id ? 'var(--color-border-primary)' : 'var(--color-border-tertiary)'}`,
              background: activeSP === sp.id ? 'var(--color-background-secondary)' : 'transparent',
              fontSize: '12px',
              cursor: 'pointer',
              color: activeSP === sp.id ? 'var(--color-text-primary)' : 'var(--color-text-secondary)',
              whiteSpace: 'nowrap',
              fontWeight: activeSP === sp.id ? 500 : 400,
            }}
          >
            {sp.name}
          </button>
        ))}
        <button
          onClick={() => { setEditingNames(salespeople.map(s => s.name)); setShowManageTeam(true); }}
          style={{ padding: '4px 10px', borderRadius: '20px', border: '0.5px solid var(--color-border-tertiary)', background: 'transparent', fontSize: '12px', cursor: 'pointer', color: 'var(--color-text-secondary)', marginLeft: '4px' }}
        >
          ✎ Edit Names
        </button>
      </div>

      <div style={outerGridStyle}>
        {[
          { label: 'total contacts', val: allContacts.length, color: 'var(--color-text-primary)' },
          { label: 'follow ups today', val: todayFollowUps.length, color: todayFollowUps.length > 0 ? '#EF9F27' : 'var(--color-text-primary)' },
          { label: 'interested', val: (stats['spoke - interested'] || 0) + (stats['follow up'] || 0), color: '#1D9E75' },
          { label: 'deals closed', val: stats['deal closed'] || 0, color: '#0F6E56' },
          { label: 'not called', val: stats['not called'] || 0, color: 'var(--color-text-secondary)' },
        ].map((s) => (
          <div key={s.label} style={{ background: 'var(--color-background-secondary)', borderRadius: 'var(--border-radius-md)', padding: '12px 14px' }}>
            <div style={{ fontSize: '13px', color: 'var(--color-text-secondary)', marginBottom: '4px' }}>{s.label}</div>
            <div style={{ fontSize: '22px', fontWeight: 500, color: s.color }}>{s.val}</div>
          </div>
        ))}
      </div>

      {todayFollowUps.length > 0 && (
        <div style={{ background: 'var(--color-background-warning)', border: '0.5px solid var(--color-border-warning)', borderRadius: 'var(--border-radius-md)', padding: '10px 14px', marginBottom: '1rem', fontSize: '13px', color: 'var(--color-text-warning)' }}>
          <span aria-hidden="true" style={{ marginRight: '6px' }}>🔔</span>
          <strong>Follow ups due today:</strong> {todayFollowUps.map((c) => `${c.first} ${c.last}`).join(', ')}
        </div>
      )}

      <div style={{ display: 'flex', gap: '6px', marginBottom: '1rem', borderBottom: '0.5px solid var(--color-border-tertiary)', paddingBottom: '8px', alignItems: 'center' }}>
        {[["list", 'contacts'], ["followups", 'follow-ups'], ["pipeline", 'pipeline'], ["reports", 'reports'], ["battle", 'battle stats']].map(([v, label]) => (
          <button
            key={v}
            onClick={() => setView(v)}
            style={{
              padding: '6px 14px',
              borderRadius: 'var(--border-radius-md)',
              border: `0.5px solid ${view === v ? 'var(--color-border-primary)' : 'var(--color-border-tertiary)'}`,
              background: view === v ? 'var(--color-background-secondary)' : 'transparent',
              fontSize: '13px',
              cursor: 'pointer',
              color: 'var(--color-text-primary)',
            }}
          >
            {label}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleExcelImport} />
        <button
          onClick={() => fileInputRef.current?.click()}
          style={{ padding: '6px 14px', borderRadius: 'var(--border-radius-md)', border: '0.5px solid var(--color-border-tertiary)', background: 'transparent', fontSize: '13px', cursor: 'pointer', color: 'var(--color-text-primary)' }}
        >
          + Import Excel
        </button>
        {importResult && <span style={{ fontSize: '12px', color: importResult.startsWith('✓') ? '#1D9E75' : '#E24B4A' }}>{importResult}</span>}
      </div>

      {view === 'pipeline' && (
        <div>
          <h3 style={{ fontSize: '16px', fontWeight: 500, marginBottom: '1rem' }}>Pipeline overview</h3>
          {STATUSES.map((s) => {
            const count = stats[s] || 0;
            const pct = Math.round((count / allContacts.length) * 100);
            return (
              <div key={s} style={{ marginBottom: '12px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '4px' }}>
                  <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ width: '10px', height: '10px', borderRadius: '50%', background: STATUS_COLORS[s], display: 'inline-block' }}></span>
                    {s}
                  </span>
                  <span style={{ color: 'var(--color-text-secondary)' }}>{count}</span>
                </div>
                <div style={{ height: '6px', background: 'var(--color-background-secondary)', borderRadius: '3px', overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${pct}%`, background: STATUS_COLORS[s], borderRadius: '3px' }}></div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {view === 'battle' && (
        <div>
          <h3 style={{ fontSize: '16px', fontWeight: 500, marginBottom: '4px' }}>Battle Stats — All Salespeople</h3>
          <p style={{ fontSize: '12px', color: 'var(--color-text-secondary)', marginBottom: '1.25rem' }}>All logged activity per salesperson — every status counts</p>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
              <thead>
                <tr style={{ borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                  <th style={{ textAlign: 'left', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400, width: '32px' }}>#</th>
                  <th style={{ textAlign: 'left', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400 }}>Name</th>
                  <th style={{ textAlign: 'center', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400, whiteSpace: 'nowrap' }}>Today</th>
                  <th style={{ textAlign: 'center', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400, whiteSpace: 'nowrap' }}>This Week</th>
                  <th style={{ textAlign: 'center', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400, whiteSpace: 'nowrap' }}>Last Week</th>
                  <th style={{ textAlign: 'center', padding: '8px 10px', color: 'var(--color-text-secondary)', fontWeight: 400, whiteSpace: 'nowrap' }}>This Month</th>
                </tr>
              </thead>
              <tbody>
                {battleStats.map((sp, i) => {
                  const isMe = sp.id === activeSP;
                  const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : '';
                  const maxToday = Math.max(...battleStats.map(s => s.today), 1);
                  const maxWeek = Math.max(...battleStats.map(s => s.thisWeek), 1);
                  const maxLastWeek = Math.max(...battleStats.map(s => s.lastWeek), 1);
                  const maxMonth = Math.max(...battleStats.map(s => s.thisMonth), 1);
                  return (
                    <tr
                      key={sp.id}
                      style={{
                        borderBottom: '0.5px solid var(--color-border-tertiary)',
                        background: isMe ? 'var(--color-background-secondary)' : 'transparent',
                      }}
                    >
                      <td style={{ padding: '10px 10px', color: 'var(--color-text-secondary)', fontSize: '12px' }}>
                        {medal || (i + 1)}
                      </td>
                      <td style={{ padding: '10px 10px', fontWeight: isMe ? 600 : 400 }}>
                        {sp.name}{isMe && <span style={{ fontSize: '11px', color: 'var(--color-text-secondary)', marginLeft: '6px' }}>(you)</span>}
                      </td>
                      {[
                        { val: sp.today, max: maxToday },
                        { val: sp.thisWeek, max: maxWeek },
                        { val: sp.lastWeek, max: maxLastWeek },
                        { val: sp.thisMonth, max: maxMonth },
                      ].map(({ val, max }, ci) => (
                        <td key={ci} style={{ padding: '10px 10px', textAlign: 'center' }}>
                          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '4px' }}>
                            <span style={{ fontWeight: val === max && val > 0 ? 600 : 400, color: val === max && val > 0 ? '#1D9E75' : val === 0 ? 'var(--color-text-secondary)' : 'var(--color-text-primary)' }}>
                              {val === 0 ? '—' : val}
                            </span>
                            {val > 0 && (
                              <div style={{ width: '40px', height: '3px', background: 'var(--color-background-secondary)', borderRadius: '2px', overflow: 'hidden' }}>
                                <div style={{ height: '100%', width: `${Math.round((val / max) * 100)}%`, background: val === max ? '#1D9E75' : '#378ADD', borderRadius: '2px' }} />
                              </div>
                            )}
                          </div>
                        </td>
                      ))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {view === 'reports' && (
        <div>
          <h3 style={{ fontSize: '16px', fontWeight: 500, marginBottom: '1rem' }}>Activity report</h3>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', marginBottom: '1.5rem' }}>
            {[
              { label: 'calls today', val: callStats.todayCount, color: callStats.todayCount > 0 ? '#1D9E75' : 'var(--color-text-primary)' },
              { label: 'calls this week', val: callStats.weekCount, color: 'var(--color-text-primary)' },
              { label: 'calls this month', val: callStats.monthCount, color: 'var(--color-text-primary)' },
            ].map(s => (
              <div key={s.label} style={{ background: 'var(--color-background-secondary)', borderRadius: 'var(--border-radius-md)', padding: '16px 14px', textAlign: 'center' }}>
                <div style={{ fontSize: '32px', fontWeight: 600, color: s.color }}>{s.val}</div>
                <div style={{ fontSize: '12px', color: 'var(--color-text-secondary)', marginTop: '4px' }}>{s.label}</div>
              </div>
            ))}
          </div>

          {Object.keys(callStats.monthByOutcome).length > 0 && (
            <div style={{ marginBottom: '1.5rem' }}>
              <div style={{ fontSize: '13px', fontWeight: 500, marginBottom: '10px', color: 'var(--color-text-secondary)' }}>THIS MONTH BY OUTCOME</div>
              {Object.entries(callStats.monthByOutcome).sort((a, b) => b[1] - a[1]).map(([outcome, count]) => (
                <div key={outcome} style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '8px', fontSize: '13px' }}>
                  <span style={{ width: '10px', height: '10px', borderRadius: '50%', background: STATUS_COLORS[outcome] || '#888', flexShrink: 0, display: 'inline-block' }} />
                  <span style={{ flex: 1 }}>{outcome}</span>
                  <span style={{ fontWeight: 500 }}>{count}</span>
                </div>
              ))}
            </div>
          )}

          <div>
            <div style={{ fontSize: '13px', fontWeight: 500, marginBottom: '10px', color: 'var(--color-text-secondary)' }}>TODAY'S CALLS</div>
            {callStats.todayCalls.length === 0 && (
              <p style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>No calls logged today.</p>
            )}
            {callStats.todayCalls.map((call, i) => (
              <div key={i} style={{ padding: '10px 0', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '2px' }}>
                  <span style={{ fontWeight: 500 }}>{call.name}</span>
                  <span style={{ color: 'var(--color-text-secondary)', fontSize: '12px' }}>{call.time}</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                  <span style={{ color: 'var(--color-text-secondary)' }}>{call.note || '—'}</span>
                  <span style={{ color: STATUS_COLORS[call.status] || 'var(--color-text-secondary)' }}>{call.status}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {view === 'followups' && (
        <div>
          <h3 style={{ fontSize: '16px', fontWeight: 500, marginBottom: '1rem' }}>Upcoming follow-ups ({upcomingFollowUps.length})</h3>
          {upcomingFollowUps.length === 0 && <p style={{ color: 'var(--color-text-secondary)', fontSize: '14px' }}>No follow-ups scheduled.</p>}
          {upcomingFollowUps.map((c) => {
            const d = crmData[c.id] || {};
            const past = isPast(d.followUpDate);
            const today2 = isToday(d.followUpDate);
            return (
              <div
                key={c.id}
                onClick={() => openContact(c)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px',
                  padding: '10px 14px',
                  border: `0.5px solid ${today2 ? 'var(--color-border-warning)' : past ? 'var(--color-border-danger)' : 'var(--color-border-tertiary)'}`,
                  borderRadius: 'var(--border-radius-md)',
                  marginBottom: '8px',
                  cursor: 'pointer',
                  background: today2 ? 'var(--color-background-warning)' : past ? 'var(--color-background-danger)' : 'var(--color-background-primary)',
                }}
              >
                <div style={{ width: '36px', height: '36px', borderRadius: '50%', background: `${companyColor(c.company)}22`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 500, color: companyColor(c.company), flexShrink: 0 }}>
                  {initials(c.first, c.last)}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: '14px', fontWeight: 500 }}>{c.first} {c.last}</div>
                  <div style={{ fontSize: '12px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{c.company}</div>
                </div>
                <div style={{ fontSize: '13px', fontWeight: 500, color: today2 ? 'var(--color-text-warning)' : past ? 'var(--color-text-danger)' : 'var(--color-text-secondary)' }}>{d.followUpDate}</div>
              </div>
            );
          })}
        </div>
      )}

      {view === 'list' && (
        <>
          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', marginBottom: '1rem' }}>
            <input
              type="text"
              placeholder="search name, company, phone..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              style={{ flex: 1, minWidth: '180px', fontSize: '13px' }}
            />
            <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} style={{ fontSize: '13px', padding: '0 8px' }}>
              <option value="all">all statuses</option>
              {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
            <select value={filterCompany} onChange={(e) => setFilterCompany(e.target.value)} style={{ fontSize: '13px', padding: '0 8px', maxWidth: '180px' }}>
              <option value="all">all companies</option>
              {companies.map((co) => <option key={co} value={co}>{co}</option>)}
            </select>
            <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} style={{ fontSize: '13px', padding: '0 8px' }}>
              <option value="name">sort: name</option>
              <option value="company">sort: company</option>
              <option value="status">sort: status</option>
              <option value="followup">sort: follow-up date</option>
            </select>
          </div>

          <div style={{ fontSize: '12px', color: 'var(--color-text-secondary)', marginBottom: '10px' }}>{filtered.length} of {allContacts.length} contacts</div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
            {filtered.map((c) => {
              const d = crmData[c.id] || {};
              const status = d.status || 'not called';
              const color = STATUS_COLORS[status];
              const fuPast = isPast(d.followUpDate);
              const fuToday2 = isToday(d.followUpDate);
              return (
                <div
                  key={c.id}
                  onClick={() => openContact(c)}
                  style={cardBaseStyle}
                  onMouseEnter={(e) => { e.currentTarget.style.borderColor = 'var(--color-border-secondary)'; }}
                  onMouseLeave={(e) => { e.currentTarget.style.borderColor = 'var(--color-border-tertiary)'; }}
                >
                  <div style={{ width: '38px', height: '38px', borderRadius: '50%', background: `${companyColor(c.company)}22`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 500, color: companyColor(c.company), flexShrink: 0 }}>
                    {initials(c.first, c.last)}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: '14px', fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{c.first} {c.last}</div>
                    <div style={{ fontSize: '12px', color: 'var(--color-text-secondary)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{c.title.length > 50 ? `${c.title.slice(0, 50)}…` : c.title}</div>
                    <div style={{ fontSize: '12px', color: 'var(--color-text-secondary)' }}>{c.company} · {c.city}{c.state ? `, ${c.state}` : ''}</div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '4px', flexShrink: 0 }}>
                    <span style={{ fontSize: '11px', padding: '2px 8px', borderRadius: '20px', background: `${color}22`, color, whiteSpace: 'nowrap' }}>{status}</span>
                    {d.followUpDate && (
                      <span style={{ fontSize: '11px', color: fuToday2 ? 'var(--color-text-warning)' : fuPast ? 'var(--color-text-danger)' : 'var(--color-text-secondary)' }}>
                        📅 {d.followUpDate}
                      </span>
                    )}
                    {d.lastCalled && <span style={{ fontSize: '11px', color: 'var(--color-text-secondary)' }}>called {d.lastCalled}</span>}
                  </div>
                </div>
              );
            })}
            {filtered.length === 0 && <p style={{ color: 'var(--color-text-secondary)', fontSize: '14px', textAlign: 'center', padding: '2rem' }}>no contacts match your filters</p>}
          </div>
        </>
      )}

      {selected && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', display: 'flex', alignItems: 'flex-start', justifyContent: 'center', zIndex: 1000, padding: '20px', overflowY: 'auto' }}>
          <div style={{ background: 'var(--color-background-primary)', borderRadius: 'var(--border-radius-lg)', border: '0.5px solid var(--color-border-tertiary)', width: '100%', maxWidth: '540px', margin: 'auto' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '14px', padding: '1.25rem', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
              <div style={{ width: '48px', height: '48px', borderRadius: '50%', background: `${companyColor(selected.company)}22`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '16px', fontWeight: 500, color: companyColor(selected.company), flexShrink: 0 }}>
                {initials(selected.first, selected.last)}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: '17px', fontWeight: 500 }}>{selected.first} {selected.last}</div>
                <div style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>{selected.title.length > 60 ? `${selected.title.slice(0, 60)}…` : selected.title}</div>
                <div style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>{selected.company}</div>
              </div>
              <button onClick={() => deleteContact(selected.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-secondary)', fontSize: '13px', padding: '4px 8px' }} aria-label="delete contact" title="Delete contact">
                🗑
              </button>
              <button onClick={() => setSelected(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-secondary)', fontSize: '20px', padding: '4px' }} aria-label="close">
                ×
              </button>
            </div>
            <div style={{ padding: '1rem 1.25rem', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
              <table style={{ width: '100%', fontSize: '13px', borderCollapse: 'collapse' }}>
                {selected.phone && (
                  <tr>
                    <td style={{ color: 'var(--color-text-secondary)', padding: '3px 0', width: '24px' }}>📞</td>
                    <td style={{ padding: '3px 0' }}><a href={`tel:${selected.phone}`}>{selected.phone}</a></td>
                  </tr>
                )}
                {selected.email && (
                  <tr>
                    <td style={{ color: 'var(--color-text-secondary)', padding: '3px 0' }}>✉️</td>
                    <td style={{ padding: '3px 0' }}><a href={`mailto:${selected.email}`}>{selected.email}</a></td>
                  </tr>
                )}
                {(selected.city || selected.state) && (
                  <tr>
                    <td style={{ color: 'var(--color-text-secondary)', padding: '3px 0' }}>📍</td>
                    <td style={{ padding: '3px 0' }}>{selected.city}{selected.state ? `, ${selected.state}` : ''}</td>
                  </tr>
                )}
                {selected.linkedin && (
                  <tr>
                    <td style={{ color: 'var(--color-text-secondary)', padding: '3px 0' }}>🔗</td>
                    <td style={{ padding: '3px 0' }}><a href={selected.linkedin} target="_blank" rel="noopener noreferrer">linkedin profile</a></td>
                  </tr>
                )}
              </table>
            </div>
            <div style={{ display: 'flex', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
              {[['overview', 'overview'], ['calls', 'call log'], ['notes', 'notes']].map(([v, label]) => (
                <button
                  key={v}
                  onClick={() => setTab(v)}
                  style={{
                    flex: 1,
                    padding: '10px',
                    fontSize: '13px',
                    border: 'none',
                    cursor: 'pointer',
                    background: tab === v ? 'var(--color-background-secondary)' : 'transparent',
                    color: tab === v ? 'var(--color-text-primary)' : 'var(--color-text-secondary)',
                    borderBottom: tab === v ? '2px solid var(--color-border-primary)' : '2px solid transparent',
                    marginBottom: '-0.5px',
                  }}
                >
                  {label}
                </button>
              ))}
            </div>
            <div style={{ padding: '1rem 1.25rem' }}>
              {tab === 'overview' && (
                <div>
                  <div style={{ marginBottom: '12px' }}>
                    <label style={{ fontSize: '12px', color: 'var(--color-text-secondary)', display: 'block', marginBottom: '4px' }}>status</label>
                    <select value={editStatus} onChange={(e) => setEditStatus(e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                      {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label style={{ fontSize: '12px', color: 'var(--color-text-secondary)', display: 'block', marginBottom: '4px' }}>follow-up date</label>
                    <input type="date" value={editFollowUp} onChange={(e) => setEditFollowUp(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button onClick={saveContact} style={{ flex: 1, padding: '8px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>save changes</button>
                    <button onClick={() => { saveContact(); setSelected(null); }} style={{ padding: '8px 14px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>save & close</button>
                  </div>
                </div>
              )}
              {tab === 'calls' && (
                <div>
                  <div style={{ display: 'flex', gap: '6px', marginBottom: '10px' }}>
                    <button
                      onClick={() => logQuickCall("Wrong number", "wrong number")}
                      style={{ flex: 1, padding: '7px', fontSize: '12px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)', background: '#C47A3A22', color: '#C47A3A', border: '0.5px solid #C47A3A66' }}
                    >
                      Wrong Number
                    </button>
                    <button
                      onClick={() => logQuickCall("Sent email", "sent email")}
                      style={{ flex: 1, padding: '7px', fontSize: '12px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)', background: '#378ADD22', color: '#378ADD', border: '0.5px solid #378ADD66' }}
                    >
                      Sent Email
                    </button>
                    <button
                      onClick={() => logQuickCall("Meeting set", "meeting set")}
                      style={{ flex: 1, padding: '7px', fontSize: '12px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)', background: '#9B59B622', color: '#9B59B6', border: '0.5px solid #9B59B666' }}
                    >
                      Meeting Set
                    </button>
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label style={{ fontSize: '12px', color: 'var(--color-text-secondary)', display: 'block', marginBottom: '4px' }}>log a call</label>
                    <textarea value={callNote} onChange={(e) => setCallNote(e.target.value)} placeholder="what happened on this call?" style={{ width: '100%', fontSize: '13px', height: '80px', padding: '8px', boxSizing: 'border-box', border: '0.5px solid var(--color-border-tertiary)', borderRadius: 'var(--border-radius-md)', background: 'var(--color-background-secondary)', color: 'var(--color-text-primary)' }} />
                    <div style={{ display: 'flex', gap: '8px', marginTop: '6px', alignItems: 'center' }}>
                      <select value={editStatus} onChange={(e) => setEditStatus(e.target.value)} style={{ flex: 1, fontSize: '13px' }}>
                        {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                      </select>
                      <button onClick={logCall} style={{ padding: '8px 14px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>log call ↗</button>
                    </div>
                    {editStatus === 'follow up' && (
                      <div style={{ marginTop: '8px' }}>
                        <label style={{ fontSize: '12px', color: 'var(--color-text-secondary)', display: 'block', marginBottom: '4px' }}>follow-up date</label>
                        <input type="date" value={editFollowUp} onChange={(e) => setEditFollowUp(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
                      </div>
                    )}
                  </div>
                  <div style={{ marginTop: '1rem' }}>
                    {(crmData[selected.id]?.calls || []).length === 0 && <p style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>no calls logged yet</p>}
                    {[...(crmData[selected.id]?.calls || [])].reverse().map((call, i) => (
                      <div key={i} style={{ padding: '10px 0', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', color: 'var(--color-text-secondary)', marginBottom: '4px' }}>
                          <span>{call.date} at {call.time}</span>
                          <span style={{ color: STATUS_COLORS[call.status] || 'var(--color-text-secondary)' }}>{call.status}</span>
                        </div>
                        {call.note && <div style={{ fontSize: '13px' }}>{call.note}</div>}
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {tab === 'notes' && (
                <div>
                  <label style={{ fontSize: '12px', color: 'var(--color-text-secondary)', display: 'block', marginBottom: '6px' }}>notes</label>
                  <textarea value={editNotes} onChange={(e) => setEditNotes(e.target.value)} placeholder="add notes about this contact..." style={{ width: '100%', fontSize: '13px', height: '140px', padding: '8px', boxSizing: 'border-box', border: '0.5px solid var(--color-border-tertiary)', borderRadius: 'var(--border-radius-md)', background: 'var(--color-background-secondary)', color: 'var(--color-text-primary)' }} />
                  <div style={{ display: 'flex', gap: '8px', marginTop: '8px' }}>
                    <button onClick={saveContact} style={{ flex: 1, padding: '8px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>save notes</button>
                    <button onClick={() => { saveContact(); setSelected(null); }} style={{ padding: '8px 14px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>save & close</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>

      {showManageTeam && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000, padding: '20px' }}>
          <div style={{ background: 'var(--color-background-primary)', borderRadius: 'var(--border-radius-lg)', border: '0.5px solid var(--color-border-tertiary)', width: '100%', maxWidth: '420px', maxHeight: '80vh', display: 'flex', flexDirection: 'column' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '1rem 1.25rem', borderBottom: '0.5px solid var(--color-border-tertiary)' }}>
              <span style={{ fontSize: '15px', fontWeight: 500 }}>Edit Salesperson Names</span>
              <button onClick={() => setShowManageTeam(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '20px', color: 'var(--color-text-secondary)' }}>×</button>
            </div>
            <div style={{ overflowY: 'auto', padding: '1rem 1.25rem', flex: 1, minHeight: 0 }}>
              {editingNames.map((name, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '8px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--color-text-secondary)', width: '20px', textAlign: 'right', flexShrink: 0 }}>{i + 1}</span>
                  <input
                    value={name}
                    onChange={e => {
                      const val = e.target.value;
                      setEditingNames(prev => { const u = [...prev]; u[i] = val; return u; });
                    }}
                    style={{ flex: 1, fontSize: '13px', padding: '6px 10px' }}
                    placeholder={`Salesperson ${i + 1}`}
                  />
                </div>
              ))}
            </div>
            <div style={{ padding: '1rem 1.25rem', borderTop: '0.5px solid var(--color-border-tertiary)', display: 'flex', gap: '8px' }}>
              <button onClick={() => setShowManageTeam(false)} style={{ padding: '8px 14px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)' }}>Cancel</button>
              <button
                onClick={() => {
                  const updated = salespeople.map((sp, i) => ({ ...sp, name: editingNames[i] || sp.name }));
                  setSalespeople(updated);
                  saveData(updated, SALESPEOPLE_KEY);
                  setShowManageTeam(false);
                }}
                style={{ flex: 1, padding: '8px', fontSize: '13px', cursor: 'pointer', borderRadius: 'var(--border-radius-md)', background: 'var(--color-border-primary)', color: '#fff', border: 'none' }}
              >
                Save Names
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
