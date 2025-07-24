/* Filter table by a column value

SELECT *
FROM [Volume-trafego-praca-pedagio-2024]
WHERE ([Volume-trafego-praca-pedagio-2024].[tipo_de_veiculo]='Comercial');

/* Filter table by multiple columns values

SELECT *
FROM [Volume-trafego-praca-pedagio-2024]
WHERE ((([Volume-trafego-praca-pedagio-2024].tipo_de_veiculo)='Moto') AND 
       (([Volume-trafego-praca-pedagio-2024].sentido)='Crescente'));





/* Use UCase to turn text to uppercase and make a case-insensity comparison 
you ensure that the query matches "Comercial" regardless of how it's capitalized in the database (e.g., 'comercial', 'COMERCIAL', 'Comercial' */

SELECT *
FROM [Volume-trafego-praca-pedagio-2024]
WHERE (((UCase([Volume-trafego-praca-pedagio-2024].[tipo_de_veiculo]))='Comercial'));





/* Update column tipo_de_veiculo so that its quotes " " be replaced by empty ''
/* UPDATE TableName
/* SET ColumnsNames

UPDATE [Volume-trafego-praca-pedagio-2024]
SET tipo_de_veiculo = Replace(tipo_de_veiculo, '"', '');





/* Update multiple columns
/* UPDATE TableName
/* SET ColumnsNames

UPDATE [Volume-trafego-praca-pedagio-2024]
SET tipo_de_veiculo = Replace(tipo_de_veiculo, '"', ''),
    another_column = Replace(another_column, '"', ''),
    yet_another_column = Replace(yet_another_column, '"', '');
