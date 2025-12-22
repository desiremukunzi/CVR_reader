Database schema
all tables to start with id column as primary key autoincrement,ends with timestamps(created_at,updated_at)

aircrafts
aircraft_type_id(foreign key to aircraft_types)
call_sign(string(20), unique)
empty_weight(decimal 6,2)
max_takeoff_weight(decimal 6,2)
status_id(foreign key to status)
manufactured_date(date)


aircraft_types
aircraft_category_id(foreign key to aircraft_categories)
type(string(20), unique)

aircraft_categories
name(string(20), unique)


status
name(string(20), unique)


crews
first_name(string(30))
last_name(string(30))
code(string(10), unique)
crew_type_id(foreign key to crew_types)

crew_types
name(string(20), unique)



flights
flight_date(date)
aircraft_id(foreign key to aircrafts)
crew_id(foreign key to crews)
flight_type_id(foreign key to flight_types)
checklist_type_id(foreign key to checklist_types)
departure_location(string(50))
destination_location(string(50))
takeoff_weight(decimal 6,2)
fuel_on_board_weight(decimal 6,2)
startup_time(time)
shutdown_time(time)
airborne_time(time)
landing_time(time)

flight_types
name(string(20), unique)



checklist_types
name(string(20), unique)


checklist_compliances
flight_id(foreign key to flights)
checklist_type_id(foreign key to checklist_types)
checks_not_complied(integer(3))
compliance_percentage(decimal 3,2)

checklist_items
name(string(100))
checklist_type_id(foreign key to checklist_types)


missed_checks
flight_id(foreign key to flights)
checklist_item_id(foreign key to checklist_items)

exceedances
flight_id(foreign key to flights)
parameter_id(foreign key to parameters)
number_of_exceedances(integer(3))


anomalies
flight_id(foreign key to flights)
parameter_id(foreign key to parameters)
phase_of_flight_id(foreign key to phase_of_flights)
total_anomalies(integer(5))

phase_of_flights
name(string(20), unique)

parameters
MI_17V_5_name(string(20), unique)
MI_17_1V_name(string(20), unique)
description(string(100))
binary(boolean)
aircraft_type_id