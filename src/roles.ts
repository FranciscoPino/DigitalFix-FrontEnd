export type Role = 'Admin' | 'Supervisor' | 'Cliente' | 'Auditor'

export function normalizeRole(value: unknown): Role | undefined {
	const normalized = String(value)
		.trim()
		.replace(/^role[_-]?/i, '')
		.replace(/[\s-]+/g, '_')
		.toUpperCase()
	if (normalized === 'ADMIN' || normalized === 'ADMINISTRADOR') return 'Admin'
	if (normalized === 'SUPERVISOR' || normalized === 'OPERADOR') return 'Supervisor'
	if (normalized === 'CLIENTE' || normalized === 'CUSTOMER') return 'Cliente'
	if (normalized === 'AUDITOR') return 'Auditor'
	return undefined
}
