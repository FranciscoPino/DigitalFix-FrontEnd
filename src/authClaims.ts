import { normalizeRole } from './roles'
import type { Role } from './roles'

export function readRolesFromToken(token: string | undefined): Role[] {
  if (!token) return []
  try {
    const payload = token.split('.')[1]
    const normalized = payload.replace(/-/g, '+').replace(/_/g, '/')
    const claims = JSON.parse(atob(normalized.padEnd(normalized.length + (4 - normalized.length % 4) % 4, '='))) as { roles?: unknown }
    return Array.isArray(claims.roles) ? claims.roles.map(normalizeRole).filter(isRole) : []
  } catch {
    return []
  }
}

export function readRolesFromClaims(claims: Record<string, unknown> | undefined): Role[] {
  const roles = claims?.roles
  return Array.isArray(roles) ? roles.map(normalizeRole).filter(isRole) : []
}

function isRole(value: Role | undefined): value is Role {
  return value !== undefined
}
