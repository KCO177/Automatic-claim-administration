import psycopg2

def create_claim_db():
   #>>> vzt8hnout do inputu
   conn = psycopg2.connect(
      database="postgres", user='postgres', password='kco177', host='localhost', port= '5432') #'127.0.0.1'
   conn.autocommit = True

   #Creating a cursor object using the cursor() method
   cursor = conn.cursor()

   #Preparing query to create a database
   sql = '''CREATE database claims''';

   #Creating a database
   cursor.execute(sql)
   print("Database claim_db created successfully........")

   #Closing the connection
   conn.close()


def create_table():
   #Establishing the connection
   conn = psycopg2.connect(
      database="claims", user='postgres', password='kco177', host='localhost', port= '5432'
   )
   #Creating a cursor object using the cursor() method
   cursor = conn.cursor()

   #Doping EMPLOYEE table if already exists.
   cursor.execute("DROP TABLE IF EXISTS claims")

   #Creating table as per requirement
   sql ='''CREATE TABLE claims(
      id INT, claim_date CHAR(20),
      project CHAR(20),
      part_name CHAR(20),
      part_id CHAR(20),
      customer CHAR(20),
      customer_claim_id CHAR(20),
      type_of_claim CHAR(20),
      claimed_failure CHAR(20),
      amount_of_claimed INT
   )'''
   cursor.execute(sql)
   print("Table created successfully........")
   conn.commit()
   #Closing the connection
   conn.close()

create_claim_db()
create_table()