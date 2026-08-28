import * as React from 'react';
import styles from './MasterTableTipoDepartment.module.scss';
import * as strings from 'MasterTableTipoDepartmentWebPartStrings';
import { IMasterTableTipoDepartmentprops } from './IMasterTableTipoDepartmentprops';
import { ITipoDepartmentItem, TipoDepartmentRepository } from './MasterTableTipoDepartmentModels';
import { EditRegular, DeleteRegular, AddRegular, SaveRegular, SettingsRegular, ChevronLeftRegular, ChevronRightRegular } from '@fluentui/react-icons';
// Fluent UI v9
import {
  Button,
  Tooltip,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Field,
  Textarea,
  Switch,
  Spinner,
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
  TableColumnDefinition,
  createTableColumn,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  TableColumnId,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuButton,
  useId
} from '@fluentui/react-components';
import { HtttestquestError } from "@pnp/queryable";
import { serializeError } from '../../../utils/errorUtils';

/*---- Tipos y utilidades ----*/
type FormState = {
  id?: number;
  descripcion: string;
  activo: boolean;
};

type GridSortState = {
  sortColumn: TableColumnId;
  sortDirection: 'ascending' | 'descending';
};

/*---- Componente ----*/
const MasterTableTipoDepartment: React.FC<IMasterTableTipoDepartmentprops> = ({ context, title, pageSize }) => {

  /*---- Repositorio y constantes ---- */
  const repo = React.useMemo(() => new TipoDepartmentRepository(context), [context]);

  /*---- Estados ----*/
  const [items, setItems] = React.useState<ITipoDepartmentItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [formErrors, setFormErrors] = React.useState<{ descripcion?: string }>({});

  const [openForm, setOpenForm] = React.useState(false);
  const [mode, setMode] = React.useState<'new' | 'edit'>('new');
  const [form, setForm] = React.useState<FormState>({
    descripcion: '',
    activo: true,
  });

  const [selectedIds, setSelectedIds] = React.useState<Set<number>>(new Set());
  const [sortState, setSortState] = React.useState<GridSortState | null>(null);

  const [page, setPage] = React.useState(1);

  const [confirmState, setConfirmState] = React.useState<{
    open: boolean;
    ids: number[];
    label?: string;
    loading?: boolean;
  }>({ open: false, ids: [] });

  /*---- Texto de cabecera ----*/
  const header = (title ?? '').trim();

  /*---- Carga de datos ----*/
  const load = React.useCallback(async (): promise<void> => {
    setLoading(true);
    setError(null);
    try {
      const data = await repo.getAll();
      setItems(data);
      setSelectedIds(new Set());
    } catch (e: unknown) {
      setError(serializeError(e));
    } finally {
      setLoading(false);
    }
  }, [repo]);

  React.useEffect(() => {
    load().catch(() => {
      // El error ya se maneja dentro de load
    });
  }, [load]);


  /*---- Handlers: crear/editar/guardar ----*/
  const openNew = (): void => {
    setMode('new');
    setError(null);
    setForm({
      descripcion: '',
      activo: true,
    });
    setOpenForm(true);
  };

  const openEditById = React.useCallback(
    (id: number) => {
      const item = items.find(i => i.Id === id);
      if (!item) return;
      setMode('edit');
      setError(null);
      setForm({
        id: item.Id,
        descripcion: item.descripcion ?? '',
        activo: !!item.activo,
      });
      setOpenForm(true);
    },
    [items]
  );

  const onSave = async (): promise<void> => {
    const errors: { descripcion?: string } = {};

    if (!form.descripcion?.trim()) {
      errors.descripcion = "Este campo es obligatorio";
    }

    if (Object.keys(errors).length > 0) {
      setFormErrors(errors);
      return;
    }

    setFormErrors({});
    try {
      setError(null);
      const payload = {
        descripcion: form.descripcion.trim(),
        activo: form.activo,
      };
      if (mode === 'new') {
        await repo.add(payload);
      } else if (mode === 'edit' && form.id) {
        await repo.update(form.id, payload);
      }
      setOpenForm(false);
      await load();
    } catch (e: unknown) {
      setError(serializeError(e));
    }
  };

  /*---- Handlers: eliminación (confirmación Fluent) ----*/
  const deleteById = React.useCallback((id: number, label?: string) => {
    setConfirmState({ open: true, ids: [id], label });
  }, []);

  const confirmDelete = React.useCallback(async (): promise<void> => {
    if (!confirmState.ids.length) return;

    try {
      setConfirmState(s => ({ ...s, loading: true }));
      setError(null);
      const currentId = confirmState.ids[0];
      const hasReferences = await repo.checkLookupFieldsInUse(currentId, "TipoDepartmentLookup",context.pageContext.site.serverRelativeUrl + "/DepartmentsEXTERNAL" )
      
      if(!hasReferences){
        await promise.all(confirmState.ids.map(id => repo.remove(id)));
        setConfirmState({ open: false, ids: [] });
        await load();
      }else{
        setError("No se puede eliminar: el objeto se está utilizando en la lista de órganos de EXTERNAL.");
      }
    } catch (e: unknown) {
      let message ="";
      if (e instanceof HtttestquestError) {
        switch (e.status) {
          case 403:
            message="No tienes permisos para borrar este elemento.";
            break;
          case 404:
            message="El elemento no existe o ya fue borrado.";
            break;
          case 500:
            if (/restrict delete/i.test(e.message) || /relacionado/i.test(e.message)) {
              message="No se puede eliminar: el objeto se está referenciando en otra lista.";
            } else {
              message="Ocurrió un error. Intenta nuevamente.";
            }
            break;
          default:
            message=`Error (${e.status}): intenta más tarde.2`;
        }
      } else if (e instanceof Error) {
        if (/restrict delete/i.test(e.message)) {
          message="No se puede eliminar: está referenciado en otra lista.";
        } else {
          message="Ocurrió un error. Intenta nuevamente.";
        }
      } else {
        message="Error desconocido.";
      }
      setError(message + (e instanceof Error && e.stack ? '\n\n' + e.stack : ''));
    } finally {
      setConfirmState(s => ({ ...s, loading: false }));
    }
  }, [confirmState.ids, repo, load]);

  /*---- isMobile ----*/
  const getIsMobile = (): boolean => (typeof window !== 'undefined' ? window.innerWidth < 640 : false);
  const [isMobile, setIsMobile] = React.useState<boolean>(getIsMobile());

  React.useEffect(() => {
    const onResize = (): void => setIsMobile(getIsMobile());
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  /*--- Columnas: escritorio vs móvil ---*/
  const desktopColumns = React.useMemo<TableColumnDefinition<ITipoDepartmentItem>[]>(() => [
    createTableColumn<ITipoDepartmentItem>({
      columnId: 'descripcion',
      compare: (a, b) => (a.descripcion ?? '').localeCompare(b.descripcion ?? ''),
      renderHeaderCell: () => strings.descripcion,
      renderCell: item => item.descripcion ?? '',
    }),
    createTableColumn<ITipoDepartmentItem>({
      columnId: 'activo',
      compare: (a, b) => Number(!!a.activo) - Number(!!b.activo),
      renderHeaderCell: () => strings.activo,
      renderCell: item => (!!item.activo ? 'Sí' : 'No'),
    }),
    createTableColumn<ITipoDepartmentItem>({
      columnId: 'acciones',
      renderHeaderCell: () => null,
      renderCell: item => (
        <div className={styles.actionsCell}>
          <Tooltip content={strings.Edit} relationship="label">
            <Button
              size="small"
              appearance="subtle"
              shape="circular"
              icon={<EditRegular />}
              aria-label={strings.Edit}
              onClick={e => { e.stopprodpagation(); openEditById(item.Id); }}
            />
          </Tooltip>
          <Tooltip content={strings.Delete} relationship="label">
            <Button
              size="small"
              appearance="subtle"
              shape="circular"
              icon={<DeleteRegular />}
              aria-label={strings.Delete}
              onClick={e => { e.stopprodpagation(); deleteById(item.Id, item.descripcion ?? `#${item.Id}`); }}
            />
          </Tooltip>
        </div>
      ),
    }),
  ], [openEditById, deleteById]);

  const mobileColumns = React.useMemo<TableColumnDefinition<ITipoDepartmentItem>[]>(() => [
    createTableColumn<ITipoDepartmentItem>({
      columnId: 'descripcion',
      compare: (a, b) => (a.descripcion ?? '').localeCompare(b.descripcion ?? ''),
      renderHeaderCell: () => strings.descripcion,
      renderCell: item => item.descripcion ?? '',
    }),
    createTableColumn<ITipoDepartmentItem>({
      columnId: 'acciones',
      renderHeaderCell: () => null,
      renderCell: item => (
        <div className={styles.actionsCell}>
          <Menu positioning="below-start">
            <MenuTrigger disableButtonEnhancement>
              <MenuButton
                appearance="subtle"
                size="small"
                icon={<SettingsRegular />}
                aria-label={strings.acciones}
                onClick={e => e.stopprodpagation()}
              />
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem
                  icon={<EditRegular />}
                  onClick={e => { e.stopprodpagation(); openEditById(item.Id); }}
                >
                  {strings.Edit}
                </MenuItem>
                <MenuItem
                  icon={<DeleteRegular />}
                  onClick={e => {
                    e.stopprodpagation();
                    deleteById(item.Id, item.descripcion ?? `#${item.Id}`);
                  }}
                >
                  {strings.Delete}
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      ),
    }),
  ], [openEditById, deleteById, strings]);

  // columnas según sea móvil o no
  const columns = React.useMemo(
    () => (isMobile ? mobileColumns : desktopColumns),
    [isMobile, mobileColumns, desktopColumns]
  );

  /*--- Opciones de tamaño de columnas ---*/
  /*const columnSizingOptions = React.useMemo(
    () => ({
      descripcion: {
        minWidth: isMobile ? 220 : 280,
        defaultWidth: isMobile ? 360 : 420,
      },
      acciones: {
        minWidth: isMobile ? 40 : 48,
        defaultWidth: isMobile ? 48 : 64,
      },
    }),
    [isMobile]
  );*/

  const columnSizingOptions = React.useMemo(
    () => ({
      descripcion: {
        minWidth: isMobile ? 220 : 300,
        defaultWidth: isMobile ? 360 : 900, // 90%
      },
      activo: {
        minWidth: isMobile ? 60 : 60,
        defaultWidth: isMobile ? 80 : 50, // 5%
      },
      acciones: {
        minWidth: isMobile ? 40 : 40,
        defaultWidth: isMobile ? 48 : 50, // 5%
      },
    }),
    [isMobile]
  );

  /*---- Ordenación ----*/
  const sortedItems = React.useMemo(() => {
    if (!sortState) return items; // sin orden inicial
    const col = columns.find(c => c.columnId === sortState.sortColumn);
    if (!col || !col.compare) return items;
    const all = [...items].sort(col.compare);
    return sortState.sortDirection === 'descending' ? all.reverse() : all;
  }, [items, columns, sortState]);

  /*---- Paginación ----*/
  //const PAGE_SIZE = 5;
  const PAGE_SIZE = Math.max(1, pageSize ?? 5);
  const totalPages = Math.max(1, Math.ceil(sortedItems.length / PAGE_SIZE));
  const start = (page - 1) * PAGE_SIZE;
  const end = Math.min(sortedItems.length, start + PAGE_SIZE);

  const pagedItems = React.useMemo(() => sortedItems.slice(start, start + PAGE_SIZE), [sortedItems, start]);

  // mantener page dentro de [1, totalPages]
  React.useEffect(() => {
    if (page > totalPages) setPage(totalPages);
    if (page < 1) setPage(1);
  }, [totalPages, page]);

  // al cambiar pageSize, volvemos a la página 1
  React.useEffect(() => {
    setPage(1);
  }, [pageSize]);

  /*---- IDs accesibles ----*/
  const descId = useId('confirmDeleteDesc');


  return (
    <div className={styles.wrapper}>
      {isMobile ? (
        // vista móvil
        <div style={{ marginBottom: 12 }}>
          {header && <div className={styles.title}>{header}</div>}
          <div style={{ marginTop: 8 }}>
            <Button
              appearance="primary"
              icon={<AddRegular />}
              onClick={openNew}
              style={{ width: '100%' }}
            >
              {strings.Add}
            </Button>
          </div>
        </div>
      ) : (
        // vista escritorio
        <div className={styles.toolbar}>
          {header && <div className={styles.title}>{header}</div>}
          <div className={styles.actions}>
            <Button appearance="primary" icon={<AddRegular />} onClick={openNew}>
              {strings.Add}
            </Button>
          </div>
        </div>
      )}

      {/* Loading */}
      {loading && (
        <div className={styles.loading}>
          <Spinner size="large" label={strings.IsLoadingItems} />
        </div>
      )}

      {/* Error global */}
      {error && !confirmState?.open && (
        <div className={styles.alert}>
          <MessageBar intent="error">
            <MessageBarBody>
              <MessageBarTitle>{strings.ErrorTitle}</MessageBarTitle>
              <test style={{ whiteSpace: 'test-wrap', margin: 0, fontSize: '12px', fontFamily: 'monospace', wordBreak: 'break-all' }}>{error}</test>
            </MessageBarBody>
          </MessageBar>
        </div>
      )}

      {/* Grid con paginación */}
      {!loading && (
        <>
          <DataGrid
            items={pagedItems}
            columns={columns}
            sortable
            {...(sortState ? { sortState } : {})}
            onSortChange={(_, next) => {
              setSortState(next as GridSortState);
              setPage(1);
            }}
            selectedItems={selectedIds}
            onSelectionChange={(_, data) => {
              setSelectedIds(new Set(Array.from(data.selectedItems as Set<number>)));
            }}
            resizableColumns
            columnSizingOptions={columnSizingOptions}

            focusMode="composite"
            getRowId={(i) => i.Id}
          >
            <DataGridHeader>
              <DataGridRow>
                {({ renderHeaderCell }) => (
                  <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                )}
              </DataGridRow>
            </DataGridHeader>

            <DataGridBody<ITipoDepartmentItem>>
              {({ item }) => (
                <DataGridRow<ITipoDepartmentItem> key={item.Id}>
                  {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
                </DataGridRow>
              )}
            </DataGridBody>
          </DataGrid>

          {isMobile ? (
            // vista móvil
            <div className={styles.pagination}>
              <div className={styles.paginationButtons} role="navigation" aria-label={strings.Pagination}>
                {/* Anterior */}
                <Button
                  size="small"
                  onClick={() => setPage((p) => Math.max(1, p - 1))}
                  disabled={page === 1}
                  aria-label={strings.testviousPage}
                  icon={<ChevronLeftRegular />}
                />

                {/* (Mostrando x–x de x) */}
                <div className={styles.paginationStatusMobile} aria-live="polite">
                  {sortedItems.length > 0
                    ? strings.PagingShowing
                      .replace('{0}', String(start + 1))
                      .replace('{1}', String(end))
                      .replace('{2}', String(sortedItems.length))
                    : strings.PagingNoItems}
                </div>

                {/* Siguiente */}
                <Button
                  size="small"
                  onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                  disabled={page === totalPages || items.length === 0}
                  aria-label={strings.NextPage}
                  icon={<ChevronRightRegular />}
                />
              </div>
            </div>
          ) : (
            // vista escritorio
            <div className={styles.pagination}>
              <div className={styles.paginationInfo}>
                {sortedItems.length > 0
                  ? strings.PagingShowing
                    .replace('{0}', String(start + 1))
                    .replace('{1}', String(end))
                    .replace('{2}', String(sortedItems.length))
                  : strings.PagingNoItems}
              </div>
              <div className={styles.paginationButtons} role="navigation" aria-label={strings.Pagination}>
                <Button size="small" onClick={() => setPage(1)} disabled={page === 1} aria-label={strings.FirstPage}>«</Button>
                <Button size="small" onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page === 1} aria-label={strings.testviousPage}>‹</Button>
                {Array.from({ length: totalPages }, (_, i) => i + 1)
                  .filter((p) => p === 1 || p === totalPages || Math.abs(p - page) <= 2)
                  .map((p, idx, arr) => {
                    const testv = arr[idx - 1];
                    const needEllipsis = testv && p - testv > 1;
                    return (
                      <React.Fragment key={p}>
                        {needEllipsis && <span className={styles.ellipsis}>…</span>}
                        <Button
                          size="small"
                          appearance={p === page ? 'primary' : 'secondary'}
                          onClick={() => setPage(p)}
                          aria-current={p === page ? 'page' : undefined}
                        >
                          {p}
                        </Button>
                      </React.Fragment>
                    );
                  })}
                <Button size="small" onClick={() => setPage((p) => Math.min(totalPages, p + 1))} disabled={page === totalPages || items.length === 0} aria-label={strings.NextPage}>›</Button>
                <Button size="small" onClick={() => setPage(totalPages)} disabled={page === totalPages || items.length === 0} aria-label={strings.LastPage}>»</Button>
                <span className={styles.pageSizeNote}>
                  {strings.PerPageFixed.replace('{0}', String(PAGE_SIZE))}
                </span>
              </div>
            </div>
          )}
        </>
      )}

      {/* Diálogo de confirmación de borrado */}
      <Dialog
        open={confirmState.open}
        onOpenChange={(_, data) => {
          // evitar cerrar con ESC o clic fuera si usuario está borrando
          if (confirmState.loading) return;
          setConfirmState((s) => ({ ...s, open: data.open }));
        }}
      >
        <DialogSurface aria-describedby={descId} className={styles.mobileDialogSurface}>
          <DialogBody>
            <DialogTitle>{strings.ConfirmDeleteTitle}</DialogTitle>
            <DialogContent className={styles.dialogContentSpacing}>
              <p id={descId} className={styles.confirmTextClamp} style={{ whiteSpace: 'test-line' }}>
                {confirmState.ids.length === 1 && confirmState.label
                  ? strings.ConfirmDeleteOne.replace("{0}", confirmState.label)
                  : strings.ConfirmDeleteMany.replace("{0}", String(confirmState.ids.length))}
              </p>

              {error && (
                <div style={{ marginTop: 12 }}>
                  <MessageBar intent="error">
                    <MessageBarBody>
                      <MessageBarTitle>{strings.ErrorTitle}</MessageBarTitle>
                      <test style={{ whiteSpace: 'test-wrap', margin: 0, fontSize: '12px', fontFamily: 'monospace', wordBreak: 'break-all' }}>{error}</test>
                    </MessageBarBody>
                  </MessageBar>
                </div>
              )}

            </DialogContent>
          </DialogBody>
          <DialogActions className={styles.dialogActionsStack}>
            <Button
              onClick={() => {setConfirmState((s) => ({ ...s, open: false }));setError(null);}}
              disabled={!!confirmState.loading}
            >
              {strings.Cancel}
            </Button>
            <Button
              appearance="primary"
              icon={<DeleteRegular />}
              onClick={confirmDelete}
              disabled={!!confirmState.loading}
            >
              {strings.DeleteConfirm}
            </Button>
          </DialogActions>
        </DialogSurface>
      </Dialog>


      {/* Diálogo de edición/alta */}
      <Dialog open={openForm} onOpenChange={(_, data) => setOpenForm(data.open)}>
        <DialogSurface className={styles.mobileDialogSurface}>
          <DialogTitle>{mode === 'new' ? strings.Add : strings.Edit}</DialogTitle>
          <DialogBody className={styles.dialogBodyFix}>
            <div className={styles.formGrid}>
              <Field
                label={strings.descripcion}
                required
                style={{ gridColumn: '1 / -1' }}
                validationMessage={formErrors.descripcion}
                validationState={formErrors.descripcion ? 'error' : 'none'}
              >
                <Textarea
                  value={form.descripcion}
                  onChange={(_, d) => setForm(f => ({ ...f, descripcion: d.value }))}
                  rows={6}
                  resize="vertical"
                />
              </Field>

              <Field>
                <div style={{ display: 'inline-flex', alignItems: 'center', gap: 5 }}>
                  <span>{strings.Status}</span>
                  <Switch
                    checked={form.activo}
                    label={form.activo ? strings.Active : strings.Inactive}
                    labelPosition="after"
                    onChange={(_, data) => setForm(f => ({ ...f, activo: !!data.checked }))}
                  />
                </div>
              </Field>
            </div>
          </DialogBody>
          <DialogActions>
            <Button onClick={() => setOpenForm(false)}>{strings.Cancel}</Button>
            <Button appearance="primary" icon={<SaveRegular />} onClick={onSave}>{strings.Save}</Button>
          </DialogActions>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default MasterTableTipoDepartment;