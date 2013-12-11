import pyodbc
import sys
import wx

import datetime as dt

import General as gn

eng04_connection = None



def connect_to_eng04_database():
	connection_string = "DSN=eng04_sql"
	db_connection = pyodbc.connect(connection_string)
	return db_connection


#return list of query results given an SQL statement
def query(sql, connection=None, commit=False):
	#make the eng04 db default if none specified
	if connection == None:
		connection = eng04_connection
		
	cursor = connection.cursor()
	result = cursor.execute(sql).fetchall()
	if result:
		if len(result[0]) == 1:
			if commit:
				connection.commit()
				
			return list(zip(*result)[0])
	
	if commit:
		connection.commit()
	
	return result


def edit(sql, connection=None):
	#make the eng04 db default if none specified
	if connection == None:
		connection = eng04_connection
		
	cursor = connection.cursor()
	cursor.execute(sql)
	
	connection.commit()


###def update(update_table, field_value_pairs, update_where, record_changes_table):
## this was going to be a generic function like the others but that went out the window...
def update_fsr(field_value_pairs, fsr_id, who_changed, when_changed):
	connection = eng04_connection
	
	#first, get the previous values so we can see which ones changed...
	sql = "SELECT "
	sql += ", ".join(zip(*field_value_pairs)[0])
	sql += " FROM tss.fsrs WHERE id={}".format(fsr_id)
	prev_values = query(sql)[0]

	for index, field_value_pair in enumerate(field_value_pairs):
		field = field_value_pair[0]
		new_value = field_value_pair[1]
		prev_value = prev_values[index]
		
		#so we don't record changes from NULL to empty string
		#if prev_value == None:
		#	prev_value = ''
		if (new_value == '' and prev_value == None) or (new_value == None and prev_value == ''):
			continue
		
		#record any changes in fields
		if new_value != prev_value:
			insert('tss.field_changes', (
				("table_name", 'tss.fsrs'), 
				("table_id", fsr_id),
				("field", field),
				("previous_value", prev_value),
				("new_value", new_value),
				("who_changed", who_changed),
				("when_changed", when_changed),)
			)

	#now update the record
	sql = "UPDATE tss.fsrs SET "
	
	for field_value_pair in field_value_pairs:
		field = field_value_pair[0]
		value = field_value_pair[1]

		if isinstance(value, bool):
			if value:
				sql += "{}=1, ".format(field)
			else:
				sql += "{}=0, ".format(field)
		elif isinstance(value, (int, long, float, complex)) or value == None:
			if value == None:
				value = 'NULL'
			sql += "{}={}, ".format(field, value)
		else:
			sql += "{}='{}', ".format(field, gn.clean(str(value)))

	sql = sql[:-2]
	sql += " WHERE id={}".format(fsr_id)
	
	edit(sql)


def update_rsr(field_value_pairs, rsr_id, who_changed, when_changed):
	connection = eng04_connection
	
	#first, get the previous values so we can see which ones changed...
	sql = "SELECT "
	sql += ", ".join(zip(*field_value_pairs)[0])
	sql += " FROM tss.rsrs WHERE id={}".format(rsr_id)
	prev_values = query(sql)[0]

	for index, field_value_pair in enumerate(field_value_pairs):
		field = field_value_pair[0]
		new_value = field_value_pair[1]
		prev_value = prev_values[index]
		
		#convert new date in string format to datetime format for comparison
		if isinstance(prev_value, dt.datetime):
			try:
				new_value = dt.datetime.strptime(new_value, '%m/%d/%Y')
			except:
				pass
		
		#so we don't record changes from NULL to empty string
		#if prev_value == None:
		#	prev_value = ''
		if (new_value == '' and prev_value == None) or (new_value == None and prev_value == ''):
			continue
		
		#record any changes in fields
		if new_value != prev_value:
			insert('tss.field_changes', (
				("table_name", 'tss.rsrs'), 
				("table_id", rsr_id),
				("field", field),
				("previous_value", prev_value),
				("new_value", new_value),
				("who_changed", who_changed),
				("when_changed", when_changed),)
			)

	#now update the record
	sql = "UPDATE tss.rsrs SET "
	
	
	for field_value_pair in field_value_pairs:
		field = field_value_pair[0]
		value = field_value_pair[1]

		if isinstance(value, bool):
			if value:
				sql += "{}=1, ".format(field)
			else:
				sql += "{}=0, ".format(field)
		elif isinstance(value, (int, long, float, complex)) or value == None:
			if value == None:
				value = 'NULL'
			sql += "{}={}, ".format(field, value)
		else:
			sql += "{}='{}', ".format(field, gn.clean(str(value)))

	sql = sql[:-2]
	sql += " WHERE id={}".format(rsr_id)
	
	edit(sql)
	
	

def insert(table, field_value_pairs, connection=None):
	#make the eng04 db default if none specified
	if connection == None:
		connection = eng04_connection
	
	sql = "INSERT INTO {} (".format(table)
	
	for field_value_pair in field_value_pairs:
		field = field_value_pair[0]
		value = field_value_pair[1]
		
		if value != '' and value != '...':
			sql += "{}, ".format(field)
			
	sql = sql[:-2]
	sql += ") OUTPUT inserted.id VALUES ("

	for field_value_pair in field_value_pairs:
		field = field_value_pair[0]
		value = field_value_pair[1]
		
		if value != '' and value != '...':
			if isinstance(value, bool):
				if value:
					sql += "1, "
				else:
					sql += "0, "
			elif isinstance(value, (int, long, float, complex)) or value == None:
				if value == None:
					value = 'NULL'
				sql += "{}, ".format(value)
			else:
				sql += "'{}', ".format(gn.clean(str(value)))

	sql = sql[:-2]
	sql += ")"
	
	cursor = connection.cursor()
	id = cursor.execute(sql).fetchone()[0]
	connection.commit()
	
	return id
	

def get_table_column_names(table, presentable=False):
	cursor = eng04_connection.cursor()

	column_names = zip(*cursor.execute("SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='{}' AND TABLE_NAME='{}' ORDER by ORDINAL_POSITION".format(table.split('.')[0], table.split('.')[1])).fetchall())[3]
		
	if presentable:
		presentable_column_names = []
		
		for column_name in column_names:
			presentable_name = column_name
			
			#uppercase the first letter
			presentable_name = presentable_name[0].upper() + presentable_name[1:]
			
			#uppercase any letter after an underscore and replace underscore with space
			while 1:
				underscore_index = presentable_name.find('_')
				if underscore_index == -1:
					break
				
				#uppercase letter after underscore
				presentable_name = presentable_name[:underscore_index+1] + \
								 presentable_name[underscore_index+1].upper() + \
								 presentable_name[underscore_index+2:]

				#replace underscore with space
				presentable_name = presentable_name.replace('_', ' ', 1)

			presentable_column_names.append(presentable_name)
		
		return presentable_column_names
	else:
		return column_names


#this function will run in it's own thread so that a slow query won't lock_held
# the GUI up. Pass it the SQL query and function you want the results to go to
def query_one_threaded(sql, reference_number_used, function_to_pass_results):
	threaded_connection = connect_to_eng04_database()
	cursor = threaded_connection.cursor()
	result = cursor.execute(sql).fetchone()
	threaded_connection.close()
	function_to_pass_results(reference_number_used, result)


#this function will run in it's own thread so that a slow query won't lock_held
# the GUI up. Pass it the SQL query and function you want the results to go to
def query_multy_threaded(sql, reference_number_used, function_to_pass_results):
	threaded_connection = connect_to_eng04_database()
	cursor = threaded_connection.cursor()
	result = cursor.execute(sql).fetchall()
	threaded_connection.close()
	function_to_pass_results(reference_number_used, result)

