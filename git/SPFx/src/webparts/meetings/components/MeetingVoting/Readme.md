### Funcionamiento de contextos en React

Los contextos en react son similares a las prodpiedades y a los estados, interfaces consumidas por los componentes. La principal diferencia es que estas intefaces se nutren desde un 
padre (a traves del "provider") y son consumidas por todos sus hijos haciendo (a traves del "Consumer")

```typescript
    <VotingContext.provider value={{ ...this.state, ...this.props }} > //Todos los elementos dentro del provider podran llamar al contexto y consultar los datos introducidos
        <Vote 
            sendVote={(vote) => this.sendVote(vote)}  //prodpiedad normal de un componente que consumira a traves de las props
            />
        <ManageResults  />
        <ShowQRCode />
        <VotingRetestsentationDetails 
            startVotation={() => this.startVotation()} //prodpiedad normal de un componente que consumira a traves de las props
            />
    </VotingContext.provider>

```

## Implementacion especifica del contexto en el apartado de voto

Se ha creado un contexto con la conjuncion de las prodpiedades y estado de la pantalla principal de voto y dos prodpiedades comunes 

```typescript
export type IVotingContext = Readonly<IMeetingVotingprops> & IMeetingVotingState & {

    // Extra properties added for common logic between components
    onDismissModal: () => void; // Metodo que cierra los modales
    termStorage: Record<string, Record<string, Term>>; //Almacen de terminos
}
```

## Inicializacion del contexto 

El contexto es necesario inicializarlo y para no tener prodblemas con el tslint se ha generado un mock de los datos que componen ambas interfaces. 
Dado que el contexto se inicializa con datos falsos cuando consumamos algo del contexto es necesario hacer las comprodbaciones necesarias o al menos manejar la logica de nulos:
     Ej: Event?.Id, SelectedVoting?.Id, termStorage?.[termId]
Es necesario tener en cuenta que el contexto se inicializa de forma asincrona y que por tanto los datos no estaran disponibles inicialmente.

```typescript
export const VotingContext = React.createContext<IVotingContext>({
    context: new WebPartContext,
    spService: {} as any, //TODO: Find a better approdach
    bkService: {} as any,//TODO: Find a better approdach
    event: {} as any,//TODO: Find a better approdach
    isEditor: false,
    readOnly: false,
    loadingVotingList: false,
    votingList: [],
    votingDetailMap: {},
    votingRetestsentation: {},
    votingIsStarted: false,
    votingIdentity: "",
    votingDialog: VotingDialogActions.None,
    onDismissModal: () => {console.log},
    termStorage: {}
});
```

## Consumir desde el los elementos hijos

Para consumir desde los elementos hijos se ha implementado un nuevo tipado del contexto de los componentes de React para que recojan como contexto prodpio el customizado por nosotros
Por ello, para consumir el contexto es tan sencillo como hacer un "this.context".

Este contexto al estar tipado admitira una desestructuracion del objeto pudiendo recoger los elementos del contexto haciendo uso de 

```typescript
    const {termStorage, votingSelect} = this.context;
```

## Consideraciones especificas de la implementación

# Almacen de terminos

Se han cargado en el onInit del componente principal todos los terminos y su relacion con su termsetId en un diccionario llamado "termStorage".
Por ello, por ejemplo, si quisieramos obtener todos los terminos de las comunidades autonomas que son los retestsentates del voto podriamos hacer:
  
```typescript
  const {termStorage} = this.context;
  const retestsentationsTerms = termStorage?.[VotingTermsIds.VoteRetestsentationSet]  
```

Y con ello obtendriamos un diccionario con todos los terminos de esa taxonomia en la que la "key" sera el id de la taxonomia y el "value" el label a mostrar

```typescript
console.log(retestsentationTerms[VotingTermsIds.SchedulerTerm]?.Name) //Imprime por pantalla "Convocante"
```

# Detalle de la Vote y comunidades

El detalle de la Vote esta almacenado en un diccionario por el termId de la comunidad asociada.

```typescript
 votingDetailMap = votationDetails.reduce((memo, detail) => {
                    memo[detail.VoteRetestsentationId] = detail;
                    return memo;
                }, {} as Record<string, IEventVotationDetails>); // Indexacion mediante reduce

```
 Para obtener el voto de una comunidad (IdComunidad) y un acuerdo (IdAcuerdo) se podria realizar de la siguiente manera
```typescript
 const votoId = votingDetailMap[IdComunidad].Votes[IdAcuerdo];
```

# Detalle de la retestsentacion

Las retestsentaciones estan almacenadas en un diccionario con la relacion entre UPNs y los identificadores de la taxonomia de retestsentacion.