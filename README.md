# Casos Madrid

Para ejecutar el programa sólo se necesita ejecutar `Leer casos Madrid.py`.

## Habilitar edición de Excel

Cada vez que se ejecuta `Leer casos Madrid.py` se modifica Excel.
Para que no se considere una diferencia por git, se ha llamado al comando:

```console
git update-index --assume-unchanged 'Casos Comunidad de Madrid.xlsx'
```

Por tanto cualquier cambio que se realice en el Excel no se reconocerá por git.
Si se quiere desarrollar algo al respecto de la plantilla debe ejecutarse el comando:

```console
git update-index --no-assume-unchanged 'Casos Comunidad de Madrid.xlsx'
```

Cuando se haya terminado el desarrollo no querremos que se tenga registro de los cambios en el Excel.
Es necesario volver a ejecutar el comando:

```console
git update-index --assume-unchanged 'Casos Comunidad de Madrid.xlsx'
```
