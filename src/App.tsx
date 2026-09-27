import { useEffect, useMemo, useState } from 'react'
import { InteractionRequiredAuthError } from '@azure/msal-browser'
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser'
import type { FormEvent } from 'react'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import { Activity, BarChart3, Bell, ClipboardList, Grid2X2, LogOut, Menu, Package, Search, ShieldCheck, UserRound, Wrench, X } from 'lucide-react'
import { loginRequest } from './authConfig'
import { readRolesFromClaims, readRolesFromToken } from './authClaims'
import type { Role } from './roles'
import { digitalFixApi } from './api'
import type { CatalogService, OrderStatus, SparePart, WorkOrder } from './api'
import './App.css'

type View = 'dashboard' | 'workorders' | 'catalog' | 'spareparts' | 'reports' | 'audit'
const technicians = [
  { id: 'a1000000-0000-4000-8000-000000000001', name: 'Camila Rojas' },
  { id: 'a1000000-0000-4000-8000-000000000002', name: 'Diego Muñoz' },
  { id: 'a1000000-0000-4000-8000-000000000003', name: 'Valentina Soto' },
]

const labels: Record<Role, string> = { Admin: 'Administrador', Supervisor: 'Supervisor de operaciones', Cliente: 'Cliente', Auditor: 'Auditor' }
const nav: Array<{ id: View; label: string; icon: typeof Grid2X2; roles: Role[] }> = [
  { id: 'dashboard', label: 'Resumen', icon: Grid2X2, roles: ['Admin', 'Supervisor', 'Cliente', 'Auditor'] },
  { id: 'workorders', label: 'Órdenes de trabajo', icon: ClipboardList, roles: ['Admin', 'Supervisor', 'Cliente'] },
  { id: 'catalog', label: 'Catálogo técnico', icon: Package, roles: ['Admin', 'Supervisor'] },
  { id: 'spareparts', label: 'Repuestos y stock', icon: Package, roles: ['Admin'] },
  { id: 'reports', label: 'Reportería', icon: BarChart3, roles: ['Admin'] },
  { id: 'audit', label: 'Auditoría', icon: ShieldCheck, roles: ['Admin', 'Auditor'] },
]

export default function App() {
  const { instance, accounts } = useMsal()
  const authenticated = useIsAuthenticated()
  const account = instance.getActiveAccount() || accounts[0]
  const [accessTokenRoles, setAccessTokenRoles] = useState<Role[]>([])
  const roles = useMemo(() => uniqueRoles([...readRolesFromClaims(account?.idTokenClaims as Record<string, unknown> | undefined), ...accessTokenRoles]), [account, accessTokenRoles])
  const [view, setView] = useState<View>('dashboard')
  const [mobileOpen, setMobileOpen] = useState(false)
  const visibleNav = nav.filter((item) => item.roles.some((role) => roles.includes(role)))
  const canCreateOrder = roles.some((role) => ['Admin', 'Supervisor', 'Cliente'].includes(role))

  useEffect(() => {
    if (!account) return
    let cancelled = false
    instance.acquireTokenSilent({ ...loginRequest, account })
      .then((result) => { if (!cancelled) setAccessTokenRoles(readRolesFromToken(result.accessToken)) })
      .catch((error) => {
        if (error instanceof InteractionRequiredAuthError) {
          instance.acquireTokenPopup({ ...loginRequest, account })
            .then((result) => { if (!cancelled) setAccessTokenRoles(readRolesFromToken(result.accessToken)) })
            .catch((popupError) => console.warn('No se pudo obtener consentimiento para DigitalFix', popupError))
        } else console.warn('No se pudo leer el access token de DigitalFix', error)
      })
    return () => { cancelled = true }
  }, [account?.homeAccountId, instance])

  if (!authenticated) return <Login onLogin={() => instance.loginRedirect(loginRequest)} />
  const current = nav.find((item) => item.id === view) || nav[0]
  const displayName = account?.name || account?.username || 'Usuario conectado'
  const initials = displayName.split(' ').map((part) => part[0]).slice(0, 2).join('').toUpperCase()

  if (roles.length === 0) return <AccessDenied account={account} onLogout={() => instance.logoutRedirect()} />

  return <div className="app-shell">
    <aside className={`sidebar ${mobileOpen ? 'open' : ''}`}>
      <div className="brand"><span className="brand-mark"><Wrench size={18} /></span><span><b>DigitalFix</b><small>Field operations</small></span><button className="mobile-close" onClick={() => setMobileOpen(false)}><X size={18} /></button></div>
      <div className="network"><i /> Red central</div>
      <p className="nav-caption">ESPACIO DE TRABAJO</p>
      <nav>{visibleNav.map((item) => { const Icon = item.icon; return <button className={`nav-link ${view === item.id ? 'selected' : ''}`} key={item.id} onClick={() => { setView(item.id); setMobileOpen(false) }}><Icon size={18} /><span>{item.label}</span></button> })}</nav>
      <div className="sidebar-footer"><span>OPERATIONS PLATFORM</span><span>v0.1.0 · AWS</span></div>
    </aside>
    <main className="main-area">
      <header className="topbar"><button className="mobile-menu" onClick={() => setMobileOpen(true)}><Menu size={21} /></button><div className="crumb">Operaciones <b>/</b> {current.label}</div><div className="top-actions"><button className="icon-button"><Search size={18} /></button><button className="icon-button alert"><Bell size={18} /><i /></button><div className="user-menu"><span className="avatar">{initials}</span><span><b>{displayName}</b><small>{labels[roles[0] || 'Cliente']}</small></span><button className="icon-button" onClick={() => instance.logoutRedirect()}><LogOut size={16} /></button></div></div></header>
      <section className="content"><div className="heading"><div><span className="eyebrow">SÁBADO, 26 DE SEPTIEMBRE DE 2026</span><h1>{greeting()}, {displayName.split(' ')[0]}.</h1><p>El pulso de tus operaciones de mantención.</p></div><div className="heading-actions"><span className="live"><Activity size={14} /> Datos actualizados</span>{canCreateOrder && <button className="primary" onClick={() => setView('workorders')}><ClipboardList size={16} /> Nueva orden</button>}</div></div><ViewContent view={view} roles={roles} account={account} instance={instance} onNavigate={setView} /></section>
    </main>
  </div>
}

function ViewContent({ view, roles, account, instance, onNavigate }: { view: View; roles: Role[]; account: AccountInfo; instance: IPublicClientApplication; onNavigate: (view: View) => void }) {
  if (view === 'workorders') return <WorkOrders canEdit={roles.includes('Admin') || roles.includes('Supervisor')} canCreate={roles.some((role) => ['Admin', 'Supervisor', 'Cliente'].includes(role))} account={account} instance={instance} />
  if (view === 'catalog') return <CatalogView account={account} instance={instance} canEdit={roles.includes('Admin')} />
  if (view === 'spareparts') return <SparePartsView account={account} instance={instance} />
  if (view === 'reports') return <Reports />
  if (view === 'audit') return <AuditDummy />
  return <Dashboard roles={roles} onNavigate={onNavigate} />
}

function Dashboard({ roles, onNavigate }: { roles: Role[]; onNavigate: (view: View) => void }) { return <><div className="stats">{[['Órdenes activas', '24', 'lime'], ['En desplazamiento', '08', 'amber'], ['Resolución promedio', '3h 42m', 'blue']].map(([label, value, tone]) => <article className="stat" key={label}><span className={`stat-mark ${tone}`}><Activity size={16} /></span><small>{label}</small><strong>{value}</strong><em>Datos actualizados</em></article>)}</div><div className="dashboard-grid"><section className="panel"><div className="panel-head"><div><span className="eyebrow">MONITOREO OPERACIONAL</span><h2>Órdenes que requieren atención</h2></div><button className="text-button" onClick={() => onNavigate('workorders')}>Ver todas →</button></div>{[['DF-1048', 'Falla en tablero principal', 'Providencia', 'EN DESPLAZAMIENTO', 'amber'], ['DF-1045', 'Mantención preventiva UPS', 'San Miguel', 'EN EJECUCIÓN', 'blue'], ['DF-1041', 'Revisión de canalización', 'Las Condes', 'ASIGNADA', 'purple']].map((row) => <div className="order" key={row[0]}><b>{row[0]}</b><span><strong>{row[1]}</strong><small>{row[2]}</small></span><Status label={row[3]} tone={row[4]} /><time>Hace 12 min</time></div>)}</section><section className="panel timeline"><div className="panel-head"><div><span className="eyebrow">ACTIVIDAD RECIENTE</span><h2>Timeline de la red</h2></div></div>{['Orden DF-1045 pasó a ejecución', 'Técnico asignado a DF-1048', 'Nueva orden creada por Cliente', 'Stock actualizado en catálogo'].map((text, i) => <div className="timeline-row" key={text}><i className={['lime', 'blue', 'amber', 'purple'][i]} /><span><strong>{text}</strong><small>{['Supervisor', 'Operador', 'Cliente', 'Admin'][i]} · hace {8 + i * 7} min</small></span></div>)}</section></div>{roles.includes('Admin') && <div className="insight"><BarChart3 size={20} /><span><b>La operación está dentro de objetivo.</b><small>El tiempo promedio bajó 18 minutos durante los últimos 7 días.</small></span><button className="text-button" onClick={() => onNavigate('reports')}>Abrir reportería →</button></div>}</> }

function WorkOrders({ canEdit, canCreate, account, instance }: { canEdit: boolean; canCreate: boolean; account: AccountInfo; instance: IPublicClientApplication }) {
  const [orders, setOrders] = useState<WorkOrder[]>([])
  const [services, setServices] = useState<CatalogService[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [showForm, setShowForm] = useState(false)
  const [technicianByOrder, setTechnicianByOrder] = useState<Record<string, string>>({})
  const [statusEditor, setStatusEditor] = useState<WorkOrder | null>(null)
  const [statusDraft, setStatusDraft] = useState<OrderStatus>('ASSIGNED')
  const [observationsDraft, setObservationsDraft] = useState('')
  const [form, setForm] = useState({ customerId: account.username, serviceId: '', title: '', description: '', address: '' })

  const load = () => { setLoading(true); digitalFixApi.listWorkOrders(instance, account, canEdit ? undefined : account.username).then(setOrders).catch((err) => setError(err.message)).finally(() => setLoading(false)) }
  useEffect(() => { void load(); digitalFixApi.listServices(instance, account).then(setServices).catch((err) => setError(err.message)) }, [account.homeAccountId, canEdit])
  const create = async (event: FormEvent) => { event.preventDefault(); try { const order = await digitalFixApi.createWorkOrder(instance, account, form); setOrders((current) => [order, ...current]); setShowForm(false); setForm({ customerId: account.username, serviceId: '', title: '', description: '', address: '' }); } catch (err) { setError(err instanceof Error ? err.message : 'No se pudo crear la orden') } }
  const updateStatus = async (id: string, defaultStatus: OrderStatus, technicianId?: string) => {
    const order = orders.find((item) => item.id === id)
    if (!order) return
    setStatusEditor(order)
    setStatusDraft(defaultStatus)
    setObservationsDraft('')
    setError('')
    if (technicianId) setTechnicianByOrder((current) => ({ ...current, [id]: technicianId }))
  }
  const saveStatusChange = async (event: FormEvent) => {
    event.preventDefault()
    if (!statusEditor) return
    const id = statusEditor.id
    const status = statusDraft
    const assignedTechnicianId = status === 'ASSIGNED' ? technicianByOrder[id] : undefined
    if (status === 'ASSIGNED' && !assignedTechnicianId) { setError('Selecciona un técnico para asignar la orden.'); return }
    try {
      const updated = await digitalFixApi.changeStatus(instance, account, id, status, assignedTechnicianId, observationsDraft.trim() || undefined)
      setOrders((current) => current.map((item) => item.id === id ? updated : item))
      setTechnicianByOrder((current) => { const next = { ...current }; delete next[id]; return next })
      setStatusEditor(null)
      setError('')
    } catch (err) { setError(err instanceof Error ? err.message : 'No se pudo actualizar el estado') }
  }
  return <>
    <Toolbar eyebrow="GESTIÓN DE SERVICIO" title="Órdenes de trabajo" action={canCreate ? 'Crear orden' : undefined} onAction={() => setShowForm(true)} />
    <section className="panel table-panel">
      <div className="table-top"><span>{loading ? 'Cargando órdenes...' : `${orders.length} órdenes encontradas`}</span><span className="live"><Activity size={14} /> API BFF</span></div>
      {error && <p className="api-error">{error}</p>}
      {showForm && <form className="inline-form" onSubmit={create}>
        <input required placeholder="Título de la orden" value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} />
        <select required value={form.serviceId} onChange={(event) => setForm({ ...form, serviceId: event.target.value })}><option value="">Selecciona un servicio</option>{services.map((service) => <option key={service.id} value={service.id}>{service.name}</option>)}</select>
        <input required placeholder="Dirección" value={form.address} onChange={(event) => setForm({ ...form, address: event.target.value })} />
        <input required placeholder="Descripción" value={form.description} onChange={(event) => setForm({ ...form, description: event.target.value })} />
        <button className="primary">Guardar orden</button><button type="button" className="secondary" onClick={() => setShowForm(false)}>Cancelar</button>
      </form>}
      <table><thead><tr><th>Orden</th><th>Servicio</th><th>Ubicación</th><th>Estado</th><th>Observaciones</th><th>Acción</th></tr></thead>
        <tbody>{orders.map((order) => <tr key={order.id}>
          <td><b>{order.id.slice(0, 8)}</b></td><td>{order.title}</td><td>{order.address}</td>
          <td><Status label={order.status} tone={statusTone(order.status)} /></td>
          <td className="order-observations">{order.observations || '—'}</td>
          <td>{canEdit && nextStatus(order.status) && <button className="text-button" onClick={() => updateStatus(order.id, nextStatus(order.status)!)}>Cambiar estado</button>}</td>
        </tr>)}</tbody>
      </table>
    </section>
    {statusEditor && <div className="dialog-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) setStatusEditor(null) }}>
      <section className="status-dialog" role="dialog" aria-modal="true" aria-labelledby="status-dialog-title">
        <button className="dialog-close icon-button" aria-label="Cerrar" onClick={() => setStatusEditor(null)}>×</button>
        <span className="eyebrow">ORDEN {statusEditor.id.slice(0, 8)}</span>
        <h2 id="status-dialog-title">Actualizar estado</h2>
        <p className="dialog-description">Elige una transición permitida o cancela la orden. Puedes dejar la observación vacía.</p>
        {error && <p className="api-error dialog-error">{error}</p>}
        <form className="status-form" onSubmit={saveStatusChange}>
          <label>Nuevo estado<select value={statusDraft} onChange={(event) => setStatusDraft(event.target.value as OrderStatus)}>{statusOptions(statusEditor.status).map((status) => <option key={status} value={status}>{statusLabel(status)}</option>)}</select></label>
          {statusDraft === 'ASSIGNED' && <label>Técnico<select required value={technicianByOrder[statusEditor.id] || ''} onChange={(event) => setTechnicianByOrder((current) => ({ ...current, [statusEditor.id]: event.target.value }))}><option value="">Selecciona técnico</option>{technicians.map((technician) => <option key={technician.id} value={technician.id}>{technician.name}</option>)}</select></label>}
          <label>Observaciones <span>(opcional)</span><textarea rows={4} maxLength={1000} placeholder="Agrega detalles del cambio, trabajo realizado o motivo de cancelación" value={observationsDraft} onChange={(event) => setObservationsDraft(event.target.value)} /></label>
          <div className="dialog-actions"><button type="button" className="secondary" onClick={() => setStatusEditor(null)}>Volver</button><button className="primary">Guardar cambio</button></div>
        </form>
      </section>
    </div>}
  </>
}
function CatalogView({ account, instance, canEdit }: { account: AccountInfo; instance: IPublicClientApplication; canEdit: boolean }) {
  const [services, setServices] = useState<CatalogService[]>([])
  const [showForm, setShowForm] = useState(false)
  const [error, setError] = useState('')
  const [form, setForm] = useState({ code: '', name: '', description: '', rate: 0 })
  useEffect(() => { digitalFixApi.listServices(instance, account).then(setServices).catch((err) => setError(err.message)) }, [account.homeAccountId])
  const create = async (event: FormEvent) => { event.preventDefault(); try { const service = await digitalFixApi.createService(instance, account, form); setServices((current) => [...current, service]); setShowForm(false); setForm({ code: '', name: '', description: '', rate: 0 }) } catch (err) { setError(err instanceof Error ? err.message : 'No se pudo crear el servicio') } }
  return <><Toolbar eyebrow="RECURSOS TÉCNICOS" title="Catálogo técnico" action={canEdit ? 'Nuevo servicio' : undefined} onAction={() => setShowForm(true)} /><section className="panel catalog-panel">{error && <p className="api-error">{error}</p>}{showForm && <form className="inline-form" onSubmit={create}><input required placeholder="Código" value={form.code} onChange={(event) => setForm({ ...form, code: event.target.value })} /><input required placeholder="Nombre del servicio" value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} /><input placeholder="Descripción" value={form.description} onChange={(event) => setForm({ ...form, description: event.target.value })} /><input required type="number" min="0" placeholder="Tarifa" value={form.rate} onChange={(event) => setForm({ ...form, rate: Number(event.target.value) })} /><button className="primary">Guardar servicio</button><button type="button" className="secondary" onClick={() => setShowForm(false)}>Cancelar</button></form>}<div className="catalog-grid">{services.map((service) => <article className="catalog-card" key={service.id}><span className="catalog-icon"><Package size={20} /></span><div><small>{service.code}</small><h3>{service.name}</h3><p>{service.description || 'Servicio técnico DigitalFix'}</p><strong>${service.rate.toLocaleString('es-CL')}</strong></div><Status label={service.active ? 'ACTIVO' : 'INACTIVO'} tone={service.active ? 'lime' : 'purple'} /></article>)}</div>{services.length === 0 && <p className="empty-message">No hay servicios registrados en el catálogo.</p>}</section></>
}
function SparePartsView({ account, instance }: { account: AccountInfo; instance: IPublicClientApplication }) {
  const emptyForm = { code: '', name: '', description: '', stockQuantity: 0, minimumStock: 0, unitPrice: 0 }
  const [parts, setParts] = useState<SparePart[]>([])
  const [form, setForm] = useState(emptyForm)
  const [editId, setEditId] = useState<string | null>(null)
  const [editForm, setEditForm] = useState({ name: '', description: '', stockQuantity: 0, minimumStock: 0, unitPrice: 0, active: true })
  const [showForm, setShowForm] = useState(false)
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(true)

  const load = () => { setLoading(true); digitalFixApi.listSpareParts(instance, account).then(setParts).catch((err) => setError(err.message)).finally(() => setLoading(false)) }
  useEffect(() => { void load() }, [account.homeAccountId])
  const create = async (event: FormEvent) => {
    event.preventDefault()
    try {
      const created = await digitalFixApi.createSparePart(instance, account, form)
      setParts((current) => [...current, created])
      setForm(emptyForm)
      setShowForm(false)
      setError('')
    } catch (err) { setError(err instanceof Error ? err.message : 'No se pudo crear el repuesto') }
  }
  const beginEdit = (part: SparePart) => {
    setEditId(part.id)
    setEditForm({ name: part.name, description: part.description || '', stockQuantity: part.stockQuantity, minimumStock: part.minimumStock, unitPrice: part.unitPrice, active: part.active })
  }
  const saveEdit = async () => {
    if (!editId) return
    try {
      const updated = await digitalFixApi.updateSparePart(instance, account, editId, editForm)
      setParts((current) => current.map((part) => part.id === updated.id ? updated : part))
      setEditId(null)
      setError('')
    } catch (err) { setError(err instanceof Error ? err.message : 'No se pudo actualizar el repuesto') }
  }

  return <><Toolbar eyebrow="INVENTARIO ELÉCTRICO" title="Repuestos y stock" action="Ingresar repuesto" onAction={() => setShowForm((open) => !open)} />
    <section className="panel table-panel"><div className="table-top"><span>{loading ? 'Consultando inventario...' : `${parts.length} repuestos`}</span><span className="live"><Activity size={14} /> PostgreSQL · catalog.spare_parts</span></div>
      {error && <p className="api-error">{error}</p>}
      {showForm && <form className="inline-form spare-part-form" onSubmit={create}>
        <input required placeholder="Código (ej. ELEC-CBL-01)" value={form.code} onChange={(event) => setForm({ ...form, code: event.target.value })} />
        <input required placeholder="Nombre de la pieza" value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} />
        <input placeholder="Descripción" value={form.description} onChange={(event) => setForm({ ...form, description: event.target.value })} />
        <label>Stock inicial<input required type="number" min="0" value={form.stockQuantity} onChange={(event) => setForm({ ...form, stockQuantity: Number(event.target.value) })} /></label>
        <label>Stock mínimo<input required type="number" min="0" value={form.minimumStock} onChange={(event) => setForm({ ...form, minimumStock: Number(event.target.value) })} /></label>
        <label>Precio unitario<input required type="number" min="0" step="0.01" value={form.unitPrice} onChange={(event) => setForm({ ...form, unitPrice: Number(event.target.value) })} /></label>
        <button className="primary">Guardar pieza</button><button type="button" className="secondary" onClick={() => setShowForm(false)}>Cancelar</button>
      </form>}
      <div className="spare-parts-table"><table><thead><tr><th>Código</th><th>Pieza</th><th>Stock</th><th>Mínimo</th><th>Precio</th><th>Estado</th><th>Acción</th></tr></thead><tbody>{parts.map((part) => <tr key={part.id}>
        {editId === part.id ? <><td><b>{part.code}</b></td><td><input className="table-input" value={editForm.name} onChange={(event) => setEditForm({ ...editForm, name: event.target.value })} /></td><td><input className="table-input numeric" type="number" min="0" value={editForm.stockQuantity} onChange={(event) => setEditForm({ ...editForm, stockQuantity: Number(event.target.value) })} /></td><td><input className="table-input numeric" type="number" min="0" value={editForm.minimumStock} onChange={(event) => setEditForm({ ...editForm, minimumStock: Number(event.target.value) })} /></td><td><input className="table-input numeric" type="number" min="0" step="0.01" value={editForm.unitPrice} onChange={(event) => setEditForm({ ...editForm, unitPrice: Number(event.target.value) })} /></td><td><label className="active-toggle"><input type="checkbox" checked={editForm.active} onChange={(event) => setEditForm({ ...editForm, active: event.target.checked })} /> Activo</label></td><td><div className="assign-controls"><button type="button" className="text-button" onClick={() => void saveEdit()}>Guardar</button><button type="button" className="text-button" onClick={() => setEditId(null)}>Cancelar</button></div></td></> : <><td><b>{part.code}</b></td><td><strong>{part.name}</strong><small className="part-description">{part.description}</small></td><td><span className={part.stockQuantity <= part.minimumStock ? 'stock-low' : ''}>{part.stockQuantity}</span></td><td>{part.minimumStock}</td><td>${part.unitPrice.toLocaleString('es-CL')}</td><td><Status label={part.active ? (part.stockQuantity <= part.minimumStock ? 'STOCK BAJO' : 'ACTIVO') : 'INACTIVO'} tone={part.active ? (part.stockQuantity <= part.minimumStock ? 'amber' : 'lime') : 'purple'} /></td><td><button className="text-button" onClick={() => beginEdit(part)}>Editar</button></td></>}</tr>)}</tbody></table></div>
      {!loading && parts.length === 0 && <p className="empty-message">No hay repuestos cargados en el inventario.</p>}
    </section>
  </>
}
function Reports() { return <><Toolbar eyebrow="INTELIGENCIA OPERACIONAL" title="Reportería" action="Últimos 7 días" /><div className="dummy-note"><BarChart3 size={16} /> Datos demostrativos: la reportería real se alimentará desde Kafka.</div><div className="report-grid"><section className="panel chart"><div className="panel-head"><div><span className="eyebrow">VOLUMEN DE ÓRDENES</span><h2>Órdenes por hora</h2></div><BarChart3 /></div><div className="bars">{[30, 48, 40, 68, 55, 76, 66, 92, 73, 88, 80, 100].map((height, i) => <i style={{ height: `${height}%` }} key={i} />)}</div></section><section className="panel metric"><span className="eyebrow">TIEMPO DE RESOLUCIÓN</span><strong>3h 42m</strong><em>↓ 18 min vs. período anterior</em><div className="progress"><i /></div><small>Objetivo operacional: 4h 00m</small></section></div></> }
function AuditDummy() { return <><Toolbar eyebrow="CUMPLIMIENTO" title="Auditoría de eventos" action="Filtrar eventos" /><div className="dummy-note"><ShieldCheck size={16} /> Datos demostrativos: el timeline real se alimentará desde Kafka.</div><section className="panel audit-dummy">{[['ORDER_STATUS_CHANGED', 'Orden DF-1045 pasó a EN EJECUCIÓN', 'supervisor@digitalfix.cl'], ['ORDER_ASSIGNED', 'Técnico asignado a orden DF-1048', 'operador@digitalfix.cl'], ['ORDER_CREATED', 'Nueva orden creada', 'cliente@empresa.cl']].map(([event, text, user], index) => <div className="audit-row" key={event}><span className={`audit-dot ${['lime', 'blue', 'amber'][index]}`} /><div><strong>{text}</strong><small>{user} · 26 SEP 2026 · hace {8 + index * 9} min</small></div><code>{event}</code></div>)}</section></> }
function Toolbar({ eyebrow, title, action, onAction }: { eyebrow: string; title: string; action?: string; onAction?: () => void }) { return <div className="toolbar"><div><span className="eyebrow">{eyebrow}</span><h2>{title}</h2></div>{action && <button className="primary" onClick={onAction}><Package size={16} /> {action}</button>}</div> }
function Status({ label, tone }: { label: string; tone: string }) { return <span className={`status ${tone}`}><i />{label}</span> }
function Login({ onLogin }: { onLogin: () => void }) { return <div className="login"><div className="login-card"><span className="brand-mark big"><Wrench size={26} /></span><span className="eyebrow">RED DE SERVICIO ELÉCTRICO</span><h1>Caso 6 -<br /><em>DigitalFix</em></h1><p>Gestiona órdenes, técnicos y recursos desde un solo lugar.</p><button className="primary login-button" onClick={onLogin}><UserRound size={17} /> Iniciar sesión con Microsoft</button><small>Acceso corporativo protegido por Microsoft Entra ID</small></div><div className="login-art"><span>DF / 01</span><h2>La precisión<br />también llega<br /><b>a terreno.</b></h2></div></div> }
function AccessDenied({ account, onLogout }: { account?: AccountInfo; onLogout: () => void }) { return <div className="login"><div className="login-card"><span className="brand-mark big"><ShieldCheck size={26} /></span><span className="eyebrow">ACCESO DIGITALFIX</span><h1>Rol no asignado.</h1><p>La cuenta {account?.username || 'actual'} inició sesión, pero no tiene un App Role reconocido en el token.</p><p>Asigna Admin, Supervisor, Cliente o Auditor en la Enterprise Application de Entra ID y vuelve a iniciar sesión.</p><button className="primary" onClick={onLogout}><LogOut size={16} /> Cerrar sesión</button></div><div className="login-art"><span>DF / 403</span><h2>La operación<br />necesita un<br /><b>rol válido.</b></h2></div></div> }
function greeting() { const hour = new Date().getHours(); return hour < 12 ? 'Buenos días' : hour < 19 ? 'Buenas tardes' : 'Buenas noches' }
function uniqueRoles(roles: Role[]): Role[] { return [...new Set(roles)] }
function statusTone(status: OrderStatus) { return status === 'EN_ROUTE' ? 'amber' : status === 'IN_PROGRESS' ? 'blue' : status === 'ASSIGNED' ? 'purple' : status === 'CLOSED' ? 'lime' : 'amber' }
function statusLabel(status: OrderStatus) { return ({ CREATED: 'CREATED · Creada', ASSIGNED: 'ASSIGNED · Asignada', EN_ROUTE: 'EN_ROUTE · En desplazamiento', IN_PROGRESS: 'IN_PROGRESS · En ejecución', CLOSED: 'CLOSED · Cerrada', CANCELLED: 'CANCELLED · Cancelada' })[status] }
function nextStatus(status: OrderStatus): OrderStatus | undefined { return ({ CREATED: 'ASSIGNED', ASSIGNED: 'EN_ROUTE', EN_ROUTE: 'IN_PROGRESS', IN_PROGRESS: 'CLOSED', CLOSED: 'CANCELLED' } as Partial<Record<OrderStatus, OrderStatus>>)[status] }
function statusOptions(status: OrderStatus): OrderStatus[] {
  if (status === 'CANCELLED') return []
  const next = nextStatus(status)
  const options: OrderStatus[] = [...(next ? [next] : []), 'CANCELLED']
  return [...new Set(options)]
}
