import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser'
import { loginRequest } from './authConfig'

const baseUrl = import.meta.env.VITE_API_BASE_URL || 'http://localhost:8080'
const retryPolicy = { maxAttempts: 3, baseDelayMs: 400, timeoutMs: 8000 }

export type OrderStatus = 'CREATED' | 'ASSIGNED' | 'EN_ROUTE' | 'IN_PROGRESS' | 'CLOSED' | 'CANCELLED'
export interface WorkOrder {
  id: string
  customerId: string
  serviceId: string
  technicianId?: string | null
  status: OrderStatus
  title: string
  description: string
  address: string
  createdAt: string
  observations?: string | null
}
export interface CatalogService { id: string; code: string; name: string; description?: string; rate: number; active: boolean }
export interface SparePart { id: string; code: string; name: string; description?: string; stockQuantity: number; minimumStock: number; unitPrice: number; active: boolean }

async function request<T>(instance: IPublicClientApplication, account: AccountInfo, path: string, init: RequestInit = {}): Promise<T> {
  const token = (await withTimeout(instance.acquireTokenSilent({ ...loginRequest, account }), retryPolicy.timeoutMs, 'La autenticación tardó demasiado')).accessToken
  const method = (init.method || 'GET').toUpperCase()
  const attempts = method === 'GET' || method === 'HEAD' ? retryPolicy.maxAttempts : 1

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    const controller = new AbortController()
    const timeout = window.setTimeout(() => controller.abort(), retryPolicy.timeoutMs)
    try {
      const response = await fetch(`${baseUrl}${path}`, {
        ...init,
        signal: controller.signal,
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}`, ...init.headers },
      })
      if (response.ok) return response.status === 204 ? (undefined as T) : response.json() as Promise<T>
      const responseText = await response.text()
      if (!shouldRetry(response.status) || attempt === attempts) throw new Error(`API ${response.status}: ${responseText}`)
    } catch (error) {
      if (attempt === attempts || (error instanceof Error && error.name !== 'AbortError' && !isNetworkError(error))) throw error
    } finally {
      window.clearTimeout(timeout)
    }
    await delay(retryPolicy.baseDelayMs * 2 ** (attempt - 1))
  }
  throw new Error('La solicitud agotó la política de reintentos')
}

function shouldRetry(status: number) { return status === 429 || status >= 500 }
function isNetworkError(error: Error) { return error instanceof TypeError || error.name === 'AbortError' }
function delay(milliseconds: number) { return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds)) }
function withTimeout<T>(promise: Promise<T>, milliseconds: number, message: string): Promise<T> {
  return Promise.race([promise, new Promise<T>((_, reject) => window.setTimeout(() => reject(new Error(message)), milliseconds))])
}

export const digitalFixApi = {
  listWorkOrders: (instance: IPublicClientApplication, account: AccountInfo, customerId?: string) =>
    request<WorkOrder[]>(instance, account, `/api/workorders${customerId ? `?customerId=${encodeURIComponent(customerId)}` : ''}`),
  createWorkOrder: (instance: IPublicClientApplication, account: AccountInfo, body: { customerId: string; serviceId: string; title: string; description: string; address: string }) =>
    request<WorkOrder>(instance, account, '/api/workorders', { method: 'POST', body: JSON.stringify(body) }),
  changeStatus: (instance: IPublicClientApplication, account: AccountInfo, id: string, status: OrderStatus, technicianId?: string, observations?: string) =>
    request<WorkOrder>(instance, account, `/api/workorders/${id}/status`, { method: 'PUT', body: JSON.stringify({ status, technicianId, observations }) }),
  sendToDlq: (instance: IPublicClientApplication, account: AccountInfo, queue: 'email' | 'dispatch' | 'report', mode: 'invalid-json' | 'invalid-envelope') =>
    request<{ exchange: string; routingKey: string; mode: string; expectedDlq: string }>(instance, account, `/api/workorders/demo/dlq?queue=${queue}&mode=${mode}`, { method: 'POST' }),
  listServices: (instance: IPublicClientApplication, account: AccountInfo) =>
    request<CatalogService[]>(instance, account, '/api/catalog/services'),
  createService: (instance: IPublicClientApplication, account: AccountInfo, body: { code: string; name: string; description: string; rate: number }) =>
    request<CatalogService>(instance, account, '/api/catalog/services', { method: 'POST', body: JSON.stringify(body) }),
  updateService: (instance: IPublicClientApplication, account: AccountInfo, id: string, body: { name: string; description: string; rate: number; active: boolean }) =>
    request<CatalogService>(instance, account, `/api/catalog/services/${id}`, { method: 'PUT', body: JSON.stringify(body) }),
  listSpareParts: (instance: IPublicClientApplication, account: AccountInfo) =>
    request<SparePart[]>(instance, account, '/api/catalog/spare-parts'),
  createSparePart: (instance: IPublicClientApplication, account: AccountInfo, body: { code: string; name: string; description: string; stockQuantity: number; minimumStock: number; unitPrice: number }) =>
    request<SparePart>(instance, account, '/api/catalog/spare-parts', { method: 'POST', body: JSON.stringify(body) }),
  updateSparePart: (instance: IPublicClientApplication, account: AccountInfo, id: string, body: { name: string; description: string; stockQuantity: number; minimumStock: number; unitPrice: number; active: boolean }) =>
    request<SparePart>(instance, account, `/api/catalog/spare-parts/${id}`, { method: 'PUT', body: JSON.stringify(body) }),
}
