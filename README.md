# Seguro de Vida con VBA

## Cálculo de Prima, Reserva y Estado de Resultados

Proyecto en Excel que hace uso de macros para el cálculo de probabilidades de: muerte, invalidez y cancelación de póliza. Además, calcula la prima, reservas (utilizando simulación de siniestros) y proyección de utilidad.

### Descripción

El presente es un proyecto académico elaborado de manera colaborativa para la materia de Cálculo Actuarial Avanzado. En este se utiliza VBA para el cálculo y estimación de Prima, Reserva y la elaboración del Estado de Resultados, tomando en cuenta las probabilidades de la CNSF (2013).

El proyecto abarca la elaboración de pólizas para seguros de vida y sus diferentes clasificaciones:

- Dotal
- Ordinario de Vida
- Seguro Temporal

### Parámetros de Cálculo

Los cálculos se realizan a partir de diversos parámetros:

- **Edad**: Edad del asegurado, con límite de 110 años.
- **Sexo**: Hombre o Mujer.
- **Condiciones**: Fumador o No Fumador.
- **Pago**: Mensual, Trimestral, Semestral o Anual.
- **Divisa**: USD (Dólares Estadounidenses), MXN (Pesos Mexicanos), UDIs (Unidades de Inversión).
- **Prima**: Monto que el asegurado desea pagar.
- **Suma Asegurada**.
- **Plazo**: Tiempo que se desea mantener la póliza.

Se consideran diversos criterios de riesgo, tasas y tiempo para realizar cálculos rigurosos y adaptados a cada caso.

<div align="center">
  <img src="https://github.com/user-attachments/assets/32f12aa3-a720-46db-a2ca-260abdb99540" alt="Parametros de Seguro de Vida" style="width:25%;">
</div>



### Funcionalidades

El archivo de Excel incluye cuatro botones que deben ejecutarse en orden. Además, contiene diversas hojas con funciones específicas:

#### 1. **Parámetros**

Esta hoja es la carátula principal y contiene los datos y parámetros a modificar según sea el caso.


<div align="center">
  <img width="401" alt="Screenshot 2024-12-29 at 21 07 42" src="https://github.com/user-attachments/assets/b2bedd20-4e37-4552-a14f-57c1da583958" alt="Parámetros de Seguro de Vida" style="width:25%;" />
</div>


Posteriormente se encuentran los botones para ejecutar la macro.
Al querer obtener información de una nueva póliza de debe de reiniciar el programa, posteriormente se crean los formatos, se ejecuta y por último se realiza el cálculo de escenarios.


 
- **Reserva**: Después de introducir los parámetros y ejecutar la macro, aquí aparecerá una estimación de la reserva. Se utiliza simulación Monte Carlo, ya que en una cartera existen diferentes sumas aseguradas y probabilidades de siniestros. Por ello, se simulan 1,000 siniestros considerando los lineamientos de Solvencia II.

#### 3. **Estado de Resultados (ER)**

Corresponde a los cálculos del Estado de Resultados, considerando:

- Ingresos: Primas y producto financiero.
- Egresos: Pago de siniestros, recuperacion de siniestros, caducidad y Maturity.
- Comisiones: Porcentaje que se paga a agentes y promotores.
- Gastos: Gastos de Administración y Gastos de mantenimiento.
- Reserva: Dinero que se aparta para hacer frente a los siniestros.

<div align="center">
  <img width="373" alt="Screenshot 2024-12-29 at 21 14 59" src="https://github.com/user-attachments/assets/6f1199d3-3f8e-4e0a-b860-e83bed80cd81" alt="Ingresos ER" style="width:20%;" />
    <img width="367" alt="Screenshot 2024-12-29 at 21 17 05" src="https://github.com/user-attachments/assets/df0745b7-a713-4356-ac8a-3d5f0b31281c" alt="Egresos ER" style="width:25%;" />
  <img width="524" alt="Screenshot 2024-12-29 at 21 31 42" src="https://github.com/user-attachments/assets/6b474c2e-29d3-4bb1-b34a-eff464e4c08f" alt="Gastos ER" style="width:25%;" />
  <img width="527" alt="Screenshot 2024-12-29 at 21 32 05" src="https://github.com/user-attachments/assets/b5f02eeb-3e11-4a50-8cad-0b550f7f05c7" alt="Reserva ER" style="width:25%;" />
</div>


#### 4. **Mortalidad e Invalidez**

- **MOIn**: Esta hoja contiene tablas de Mortalidad e Invalidez (CNSF, 2013). A partir de estas tablas se calculan probabilidades de supervivencia, mortalidad, invalidez y cancelación del contrato de seguro.

<div align="center">
  <img width="641" alt="Screenshot 2024-12-29 at 21 37 27" src="https://github.com/user-attachments/assets/1aab78c3-0df1-476b-9e33-12c8b12f6d49" alt="Mortalidad e Invalidez" style="width:35%;" />
</div>


#### 5. **Comisiones y Bonos**

Incluye la compensación otorgada al agente y al promotor. Para este proyecto, se asignan de manera arbitraria.

<div align="center">
  <img width="1021" alt="Screenshot 2024-12-29 at 21 40 15" src="https://github.com/user-attachments/assets/1085409d-a76e-4989-964a-60ee6e59b133" alt="Comisiones y Bonos" style="width:80%;" />
</div>

#### 6. **Probabilidades Personales (PP)**

Esta hoja contiene el cálculo de probabilidades del asegurado, considerando los años de la póliza y sus características particulares, se presenta la probabilidad se No Muerte, la probailidad de No Invalidez y por último la propabilidad de No Cancelar la Póliza durante el tiempo de vigencia.

<div align="center">
  <img width="1021" alt="Screenshot 2024-12-29 at 21 43 04" src="https://github.com/user-attachments/assets/f5e79223-7ffb-4db2-aa98-59d3d0b13b1a" alt="Probabilidades" style="width:80%;" />
</div> 



#### 7. **Escenarios**

Genera varios escenarios para simular la siniestralidad y realiza el cálculo de:

- BEL (Best Estimate Liability)
- Percentil
- Reserva de Riesgo en Curso
- Reserva General

Se calculan 1,000 escenarios posibles.

#### 8. **AVR**

Contiene el valor futuro de la inversión de las utilidades, usando como referencia los Certificados de la Tesorería de la Federación (CETES). Para este proyecto, se utilizó la tasa de marzo de 2022.

### Instrucciones

1. Introducir los parámetros en la hoja correspondiente.
2. Reiniciar Programa
3. Ejecutar Bóton de "Dar Formato"
4. Ejecutar (Para cálculo de primas y ER)
5. Cálculo de escenarios


Este proyecto es una demostración del uso de herramientas actuariales mediante VBA en Excel, enfocándose en seguros de vida.
