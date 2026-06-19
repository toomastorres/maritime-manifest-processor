==========================================================================================================
Gastos de linea modificables:
Impo 
THC POD 100usd x CTRS (20Tns).
THC POD 120usd x CTRS (40Tns).
S/C POD 150usd x CTRS (20Tns y 40Tns) (s/c lo llamamos TOLL).
S/C POD 20usd x Tns (unicamente en carga suelte es decir no son contenedores) (s/c lo llamamos TOLL).
SWEEPING 10usd x CTRRS (20Tns y 40Tns).

Expo 
THC POL 100usd x CTRS (20Tns).
THC POL 120usd x CTRS (40Tns) .
S/C POL 150usd x CTRS (20Tns y 40Tns) (s/c lo llamamos TOLL).
S/C POL 20usd x Tns (unicamente en carga suelte es decir no son contenedores) (s/c lo llamamos TOLL).
SWEEPING 10usd x CTRRS (20Tns y 40Tns).

==========================================================================================================
Comisionable:
Basic Frt. + Open Top - TODOS LOS MONTOS COMISIONABLES INCLUIDOS EN EL BL = MONTO COMISION. Nunca puede dar negativo, en caso de hacer el calculo y este negativo el monto Comision sera 0.

Impo:
Rolling (solo cars o big vans) 1% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado.
General Cargo (carga general (roros) y carga proyecto) 2% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado.
Contenedores 2% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado y 25 USD por contenedor, calculo final es el monto mas alto entre ambos calculos.

Expo:
Rolling (solo cars o big vans) 1% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado.
General Cargo (carga general (roros) y carga proyecto) 4% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado.
Contenedores 4% sobre el Monto Comision; si el monto Comision es EUR se convierte a USD según el TC pactado y 25 USD por contenedor, calculo final es el monto mas alto entre ambos calculos.

==========================================================================================================
Reglas de Biblias:

Impo:
S/C POD tipo C (collect), si es tipo P (prepaid) dejar un mensaje en la celda aclarando. Discriminar por 20 y 150, es decir por contenedor y por carga general, agrupar los montos que sean de igual factor.
THC POD tipo C (collect), si es tipo P (prepaid) dejar un mensaje en la celda aclarando. Discriminar por 100 y 120, es decir por contenedor de 20 y 40 Tns. Agrupar los montos que sean de igual factor.
SWEEPING tipo C (collect), si es tipo P (prepaid) dejar un mensaje en la celda aclarando.
USD Y EUR PREPAID. En los Prepaid va el Monto Comision (Basic Frt. y sus variables de sumas o resta), siempre y cuando el mismo sea condicion. Discriminar montos segun tipo de Factor que corresponda, es decir si hay un Basic. Frt. de contenedores y de carga general o autos separar el monto comisionable correspondiente a cada uno de los tipos.
USD Y EUR COLLECT. En los Collect la regla es; si Basic. Frt. es tipo C, se coloca en la columna de Collect el: "Total Buenos Aires EUR" - "BAF", el BAF se descrimina en una columna, y el Monto Comision va por fuera; "Total Buenos Aires USD" - "BAF", el BAF se descrimina en una columna - S/C POD si es TIPO C - THC POD si es TIPO C, y el Monto Comision va por fuera.

Expo:
THC POL tipo P (prepaid), si es tipo C (collect) dejar un mensaje en la celda aclarando. Discriminar por 100 y 120, es decir por contenedor de 20 y 40 Tns. Agrupar los montos que sean de igual factor.
SWEEPING tipo P (prepaid), si es tipo P (collect) dejar un mensaje en la celda aclarando.
USD INCL PREPAID. "Total Buenos Aires USD" - THC POL tipo P - SWEEPING tipo P. Si esa cuenta da negativo nos da a entender que THC y SWEEPING por mas de ser tipo P, estos se pagan en MATRIZ, lo cual en la celda queda el "Total Buenos Aires" y en las celdas de THC y SWEEPING donde estaban los montos inciales se especifica la palabra "MATRIZ" que nos da la pauta de que se paga afuera y debemos recuperar.
COLLECT. "Total [Pais] [Moneda]"
ABROAD. "Total Matriz USD"

==========================================================================================================
Tipo de carga:

Identificar bien los tipos de cargas.
Los Autos correspondientes a fabricas como BIONDA, AVALON, CORVEX, GRIFON, QUASAR, PEGASO, NOVAX, ORBIS, EVORA (Premium Auto), MISTRAL, MISTRAL DELTAR, MISTRAL DELTAR - Delcar, MISTRAL DELTAR - Premiumcars deben ser identificado.
Identificar los tipos de carga proyecto como Pallets, Units, Machine, Package, Box, Case, Bundle, Unpacked, Crate, Steel, entre otros y su cantidad.
La forma de encontrar la cantidad es corroborando con lo comisionable, debe coincidir la carga con lo que se esta comisionando (osea Basic. Frt.)

==========================================================================================================