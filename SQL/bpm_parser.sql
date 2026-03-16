WITH
BPM AS (
SELECT 
	order__id,
	order__add_info__comment,
	concat(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(splitByChar('.', splitByChar('|', replace(order__add_info__comment, 'USD/сутки.', '|'))[2])[1], ' ', ''), ' ', ''), char(10), ''), '-', ': ' ), ',', concat(';', char(10))), ';') AS depo_cost,
	created_date,
	client__name,
	JSONExtractString(JSONExtractString(JSONExtractArrayRaw(order__add_info__delivery_address)[1], 'address'), 'str') AS order__add_info__delivery_address_from_address,
	state,
	--order__transport_solution__location__from__catalog_id,
	--dictGet('dict_location', 'name', order__transport_solution__location__from__catalog_id) AS order__transport_solution__location__from__catalog_name,
	--dictGet('dict_location', 'name', order__transport_solution__location__to__catalog_id) AS order__transport_solution__location__to__catalog_name,
	--
	order__transport_solution__legs,
	arrayJoin(emptyArrayToSingle(JSONExtractArrayRaw(order__transport_solution__legs))) AS legs_array,
	JSONExtractInt(legs_array,'id') AS leg_id,
	--FROM
	JSONExtractString(JSONExtractRaw(JSONExtractRaw(legs_array,'points'),'from'),'catalog_id') AS leg_points_from_catalog_id,
	dictGet('dict_location', 'name', leg_points_from_catalog_id) AS leg_points_from_catalog_name,
	JSONExtractRaw(JSONExtractString(JSONExtractRaw(legs_array,'points'),'from'),'station_id') AS leg_points_from_station_id,
	dictGet('dict_stations', 'station_name',  leg_points_from_station_id) AS leg_points_from_station_name,	
	JSONExtractString(JSONExtractRaw(JSONExtractRaw(legs_array,'points'),'to'),'catalog_id') AS leg_points_to_catalog_id,
	--TO
	JSONExtractString(JSONExtractRaw(JSONExtractRaw(legs_array,'points'),'to'),'catalog_id') AS leg_points_to_catalog_id,
	dictGet('dict_location', 'name', leg_points_to_catalog_id) AS leg_points_to_catalog_name,
	dictGet('dict_location', 'country', leg_points_to_catalog_id) AS leg_points_to_catalog_country,
	JSONExtractString(JSONExtractRaw(JSONExtractRaw(legs_array,'points'),'to'),'station_id') AS leg_points_to_station_id,
	dictGet('dict_stations', 'station_name',  leg_points_to_station_id) AS leg_points_to_station_name,	
	JSONExtractString(JSONExtractRaw(JSONExtractRaw(legs_array,'points'),'to'),'catalog_id') AS leg_points_to_catalog_id,
	--
	arrayJoin(emptyArrayToSingle(JSONExtractArrayRaw(JSONExtractRaw(legs_array,'services')))) AS services_array,
	JSONExtractInt(services_array,'id') AS service_id,
	JSONExtractString(services_array,'epu_id') AS service_epu_id,			
	arrayJoin(emptyArrayToSingle(JSONExtractArrayRaw(services_array,'details'))) AS service_details_array,
	JSONExtractInt(service_details_array,'id') AS service_details_id,
	JSONExtractString(service_details_array,'esu_id') AS service_details_esu_id,
	reverseUTF8(splitByChar('/', reverseUTF8(JSONExtractString(service_details_array, 'rate_sell_url')))[1]) AS rate_sell_url,
	JSONExtractString(service_details_array,'currency_coefficient') AS service_details_currency_coefficient,
	JSONExtractFloat(service_details_array,'cost_per_unit') AS service_details_cost_per_unit,
	JSONExtractFloat(service_details_array,'cost_per_unit_in_contract_currency') AS service_details_cost_per_unit_in_contract_currency
FROM 
	bpm__order AS BPM	-- SELECT * FROM bpm__order
WHERE
	order__id = 60020183
)
SELECT
	order__id,order__add_info__comment,depo_cost,--created_date,client__name,order__add_info__delivery_address_from_address,state,order__transport_solution__legs,legs_array,leg_id,leg_points_from_catalog_id,leg_points_from_catalog_name,leg_points_from_station_id,leg_points_from_station_name,leg_points_to_catalog_id,leg_points_to_catalog_name,leg_points_to_catalog_country,leg_points_to_station_id,leg_points_to_station_name,services_array,service_id,service_epu_id,service_details_array,service_details_id,service_details_esu_id,rate_sell_url,
	service_details_currency_coefficient--,service_details_cost_per_unit,service_details_cost_per_unit_in_contract_currency
FROM
	BPM