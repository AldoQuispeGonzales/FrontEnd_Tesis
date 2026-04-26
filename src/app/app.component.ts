import { CommonModule } from '@angular/common';
import { Component, DoCheck, HostListener, OnDestroy, OnInit } from '@angular/core';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { FormsModule, NgForm } from '@angular/forms';

declare const mapboxgl: any;

type View = 'dashboard' | 'clients' | 'orders-list' | 'orders-new' | 'orders-history' | 'routes-daily' | 'routes-assignment' | 'vehicles-list' | 'vehicles-new' | 'drivers-list' | 'drivers-new' | 'billing-boletas' | 'billing-facturas' | 'billing-credit-notes' | 'billing-sunat' | 'billing-issued' | 'reports' | 'reports-sales' | 'reports-orders' | 'reports-drivers' | 'reports-income' | 'reports-export' | 'config-company' | 'config-igv' | 'config-certificate' | 'config-series' | 'users-list' | 'users-new' | 'roles-list' | 'roles-new' | 'permissions-list' | 'permissions-edit' | 'security-audit';

type ReportsTab = 'sales' | 'orders' | 'drivers' | 'income';
type ClientType = 'empresa' | 'persona';
type DocumentType = 'RUC' | 'DNI';
type OrderStatus = 'Pendiente' | 'En ruta' | 'Entregado';
type ServiceType = 'Agua potable' | 'Piscina' | 'Obra';
type ZoneType = 'Centro' | 'Norte' | 'Sur' | 'Este' | 'Oeste';
type VehicleStatus = 'Disponible' | 'En ruta' | 'Mantenimiento';
type DriverStatus = 'Activo' | 'Inactivo' | 'Suspendido';
type BillingDocType = 'Boleta' | 'Factura' | 'Nota de crédito' | 'Nota de venta';
type SunatChannel = 'API directa SUNAT' | 'Nubefact' | 'Facturador SUNAT';
type SunatStatus = 'Borrador' | 'Aceptado' | 'Rechazado';
type UserRole = 'Admin' | 'Operador' | 'Chofer';
type UserStatus = 'Activo' | 'Inactivo' | 'Suspendido';
type PermissionType = 'read' | 'write' | 'delete' | 'admin';

type ToastKind = 'success' | 'error';

interface Toast {
  id: number;
  kind: ToastKind;
  text: string;
}

interface Client {
  id: number;
  documentType: DocumentType;
  documentNumber: string;
  name: string;
  address: string;
  phone: string;
  email: string;
  clientType: ClientType;
}

interface Order {
  id: number;
  clientId: number | null;
  clientName: string;
  clientDocument?: string;
  zone: ZoneType;
  deliveryAddress: string;
  serviceType: ServiceType;
  quantity: number | null;
  price: number | null;
  deliveryDate: string;
  status: OrderStatus;
  vehicle: string;
  driver: string;
  history: string[];
}

interface Vehicle {
  id: number;
  plate: string;
  capacity: number;
  status: VehicleStatus;
  soat: string;
  technicalReview: string;
}

interface Driver {
  id: number;
  name: string;
  dni: string;
  license: string;
  phone: string;
  status: DriverStatus;
}

interface BillingDocument {
  id: number;
  type: BillingDocType;
  clientId?: number | null;
  orderId?: number | null;
  clientDocument: string;
  clientName: string;
  detail: string;
  subtotal: number;
  igv: number;
  total: number;
  series?: string | null;
  correlative?: number | null;
  xmlPath: string;
  zipPath: string;
  pdfPath: string;
  cdrPath: string;
  certificatePath: string;
  channel: SunatChannel;
  sunatStatus: SunatStatus;
  response: string;
  createdAt?: string;
}

interface DocumentSeriesConfig {
  docType: BillingDocType;
  scope?: string;
  series: string;
  nextCorrelative: number;
}

interface SalesReport {
  date: string;
  totalSales: number;
  orderCount: number;
}

interface OrderReport {
  clientName: string;
  clientDocument: string;
  orderCount: number;
  totalAmount: number;
}

interface DriverReport {
  driverName: string;
  completedOrders: number;
  totalDeliveries: number;
  rating: number;
  onTimePercentage: number;
}

interface MonthlyIncome {
  month: string;
  income: number;
}

interface CompanyData {
  name: string;
  ruc: string;
  address: string;
  ubigeo: string;
  phone: string;
  email: string;
  urbanizacion: string;
  distrito: string;
  provincia: string;
  departamento: string;
}

// Invoice series are handled by backend.

interface CertificateConfig {
  path: string;
  password: string;
  expiration: string;
}

// SUNAT API config is handled by backend.

interface User {
  id: number;
  username: string;
  email: string;
  role: UserRole;
  status: UserStatus;
  isTemporary?: boolean;
  expiresAt?: string | null;
  temporaryHours?: number | null;
  createdAt: string;
  lastLogin: string;
  password?: string;
}

interface Role {
  id: number;
  name: UserRole;
  description: string;
  permissions: string[];
}

interface Permission {
  id: number;
  module: string;
  description: string;
  type: PermissionType;
}

interface DashboardNotification {
  id: number;
  text: string;
  createdAt: string;
  read: boolean;
  level: ToastKind;
  key?: string;
}

interface AuthAuditLog {
  id: number;
  userId: number | null;
  username: string;
  eventType: string;
  success: boolean;
  ip: string;
  userAgent: string;
  message: string;
  createdAt: string;
  details: Record<string, unknown>;
}

const DEFAULT_LOCAL_API_BASE_URL = 'http://localhost:3000/api';
const DEFAULT_REMOTE_API_BASE_URL = 'https://backend-tesis-r3zf.onrender.com/api';

function normalizeApiBaseUrl(raw: string): string {
  const base = String(raw || '').trim().replace(/\/$/, '');
  if (!base) return DEFAULT_REMOTE_API_BASE_URL;
  return base.endsWith('/api') ? base : `${base}/api`;
}

function resolveApiBaseUrl(): string {
  const configured = (window as any).__API_BASE_URL__;
  if (typeof configured === 'string' && configured.trim()) {
    return normalizeApiBaseUrl(configured);
  }

  const host = String(window.location.hostname || '').toLowerCase();
  if (host === 'localhost' || host === '127.0.0.1') return DEFAULT_LOCAL_API_BASE_URL;
  return DEFAULT_REMOTE_API_BASE_URL;
}

function toNotificationsWsUrl(apiBaseUrl: string, userId: number): string {
  const apiRoot = String(apiBaseUrl || '').replace(/\/api$/, '');
  const wsBase = apiRoot.replace(/^http/i, 'ws');
  return `${wsBase}/ws/notifications?userId=${userId}`;
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent implements DoCheck, OnInit, OnDestroy {
  constructor(private sanitizer: DomSanitizer) {}
  private readonly apiBaseUrl = resolveApiBaseUrl();
  private mapboxAccessToken = String((window as any).MAPBOX_ACCESS_TOKEN || '').trim();
  private notificationUserId = 1;
  private readonly authStorageKey = 'erp_auth_session_v1';

  private notificationsWs: WebSocket | null = null;
  private notificationsWsRetryTimer: number | null = null;
  private notificationsWsShouldRun = false;

  isAuthenticated = false;
  authUserId = 0;
  authUsername = '';
  authRole = '';
  authIsTemporary = false;
  authExpiresAt: string | null = null;
  authError = '';
  authLoading = false;
  loginUsername = '';
  loginPassword = '';
  loginPasswordVisible = false;
  loginBrandImagePath = '/assets/login/login.png';
  loginHasImage = true;
  authAccessToken = '';
  authRefreshToken = '';
  private authRefreshPromise: Promise<boolean> | null = null;

  toasts: Toast[] = [];
  private toastSeq = 0;
  private toastTimers = new Map<number, number>();

  dismissToast(id: number): void {
    const timer = this.toastTimers.get(id);
    if (timer) window.clearTimeout(timer);
    this.toastTimers.delete(id);
    this.toasts = this.toasts.filter((t) => t.id !== id);
  }

  private pushToast(kind: ToastKind, text: string, ms = 5000, notify = true): void {
    const trimmed = String(text || '').trim();
    if (!trimmed) return;

    const id = ++this.toastSeq;
    this.toasts = [{ id, kind, text: trimmed }, ...this.toasts].slice(0, 5);
    if (notify && this.isAuthenticated) this.emitDashboardNotification(trimmed, kind);

    const timer = window.setTimeout(() => this.dismissToast(id), ms);
    this.toastTimers.set(id, timer);
  }

  private transientTimers = new Map<string, number>();

  private setTransientNotice(
    messageField: string,
    errorField: string,
    kind: 'success' | 'error',
    text: string,
    ms = 5000
  ): void {
    (this as any)[messageField] = '';
    (this as any)[errorField] = '';

    const targetField = kind === 'success' ? messageField : errorField;
    (this as any)[targetField] = text;

    const prev = this.transientTimers.get(targetField);
    if (prev) window.clearTimeout(prev);

    const timer = window.setTimeout(() => {
      if ((this as any)[targetField] === text) (this as any)[targetField] = '';
      this.transientTimers.delete(targetField);
    }, ms);

    this.transientTimers.set(targetField, timer);

    // Global toast
    this.pushToast(kind, text, ms);
  }

  onLoginImageError(): void {
    this.loginHasImage = false;
  }

  private restoreAuthSession(): void {
    try {
      const raw = localStorage.getItem(this.authStorageKey);
      if (!raw) return;
      const data = JSON.parse(raw || '{}');
      const username = String(data?.username || '').trim();
      const role = String(data?.role || '').trim();
      const userId = Number(data?.userId || 0);
      const accessToken = String(data?.accessToken || '').trim();
      const refreshToken = String(data?.refreshToken || '').trim();
      const isTemporary = Boolean(data?.isTemporary);
      const expiresAt = data?.expiresAt == null || String(data?.expiresAt || '').trim() === '' ? null : String(data.expiresAt);
      if (!username) return;
      this.isAuthenticated = true;
      this.authUserId = Number.isFinite(userId) && userId > 0 ? userId : 1;
      this.authUsername = username;
      this.authRole = role;
      this.authIsTemporary = isTemporary;
      this.authExpiresAt = expiresAt;
      this.authAccessToken = accessToken;
      this.authRefreshToken = refreshToken;
      this.notificationUserId = this.authUserId;
    } catch {
      // ignore invalid storage
    }
  }

  private persistAuthSession(): void {
    try {
      localStorage.setItem(
        this.authStorageKey,
        JSON.stringify({
          username: this.authUsername,
            userId: this.authUserId,
            role: this.authRole,
            isTemporary: this.authIsTemporary,
            expiresAt: this.authExpiresAt,
            accessToken: this.authAccessToken,
            refreshToken: this.authRefreshToken,
            at: new Date().toISOString()
        })
      );
    } catch {
      // ignore
    }
  }

  private clearAuthSession(): void {
    try {
      localStorage.removeItem(this.authStorageKey);
    } catch {
      // ignore
    }
  }

  private async afterLoginBootstrap(): Promise<void> {
    this.notificationsWsShouldRun = true;
    await this.loadDashboardNotificationsFromApi();
    await this.loadOrdersFromApi();
    await this.loadBillingFromApi();
    this.connectNotificationsWs();
  }

  private async refreshAuthToken(): Promise<boolean> {
    if (!this.authRefreshToken) return false;
    if (this.authRefreshPromise) return this.authRefreshPromise;

    this.authRefreshPromise = (async () => {
      try {
        const res = await fetch(`${this.apiBaseUrl}/auth/refresh`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ refreshToken: this.authRefreshToken })
        });
        const text = await res.text();
        const data = text ? JSON.parse(text) : null;
        if (!res.ok || !data?.accessToken) return false;

        this.authAccessToken = String(data.accessToken || '');
        this.authRefreshToken = String(data.refreshToken || this.authRefreshToken);
        this.authUserId = Number(data?.user?.id || this.authUserId || 1);
        this.authUsername = String(data?.user?.username || this.authUsername);
        this.authRole = String(data?.user?.role || this.authRole);
        this.authIsTemporary = Boolean(data?.user?.isTemporary);
        this.authExpiresAt = data?.user?.expiresAt == null || String(data?.user?.expiresAt || '').trim() === '' ? null : String(data.user.expiresAt);
        this.notificationUserId = this.authUserId;
        this.persistAuthSession();
        return true;
      } catch {
        return false;
      } finally {
        this.authRefreshPromise = null;
      }
    })();

    return this.authRefreshPromise;
  }

  async loginToSystem(): Promise<void> {
    const username = String(this.loginUsername || '').trim();
    const password = String(this.loginPassword || '');
    this.authError = '';

    if (!username || !password) {
      this.authError = 'Ingresa usuario y contraseña.';
      return;
    }

    this.authLoading = true;
    try {
      const auth = await this.apiJson<any>('POST', '/auth/login', { username, password });
      this.isAuthenticated = true;
      this.authUserId = Number(auth?.user?.id || 1);
      this.authUsername = String(auth?.user?.username || username);
      this.authRole = String(auth?.user?.role || '');
      this.authIsTemporary = Boolean(auth?.user?.isTemporary);
      this.authExpiresAt = auth?.user?.expiresAt == null || String(auth?.user?.expiresAt || '').trim() === '' ? null : String(auth.user.expiresAt);
      this.authAccessToken = String(auth?.accessToken || '');
      this.authRefreshToken = String(auth?.refreshToken || '');
      this.notificationUserId = this.authUserId;
      this.loginPassword = '';
      this.persistAuthSession();
      await this.afterLoginBootstrap();
      this.pushToast('success', 'Bienvenido al sistema.', 3000, false);
    } catch (e) {
      this.authError = String((e as any)?.message || 'No se pudo iniciar sesión.');
    } finally {
      this.authLoading = false;
    }
  }

  logoutFromSystem(): void {
    const refreshToken = this.authRefreshToken;
    if (refreshToken) {
      void fetch(`${this.apiBaseUrl}/auth/logout`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ refreshToken })
      }).catch(() => undefined);
    }

    this.isAuthenticated = false;
    this.authUserId = 0;
    this.authUsername = '';
    this.authRole = '';
    this.authIsTemporary = false;
    this.authExpiresAt = null;
    this.authAccessToken = '';
    this.authRefreshToken = '';
    this.notificationUserId = 1;
    this.authError = '';
    this.loginUsername = '';
    this.loginPassword = '';
    this.dashboardNotificationsOpen = false;
    this.toasts = [];
    this.clearAuthSession();

    this.notificationsWsShouldRun = false;
    if (this.notificationsWsRetryTimer) window.clearTimeout(this.notificationsWsRetryTimer);
    this.notificationsWsRetryTimer = null;
    try {
      this.notificationsWs?.close();
    } catch {
      // ignore
    }
    this.notificationsWs = null;
  }

  ngOnInit(): void {
    void this.loadPublicRuntimeConfig();
    this.restoreAuthSession();
    if (this.isAuthenticated) {
      void this.afterLoginBootstrap();
    }
  }

  private async loadPublicRuntimeConfig(): Promise<void> {
    try {
      const res = await fetch(`${this.apiBaseUrl}/public/config`);
      if (!res.ok) return;

      const data = (await res.json()) as { mapboxPublicToken?: string | null };
      const token = String(data?.mapboxPublicToken || '').trim();
      if (token) this.mapboxAccessToken = token;
    } catch {
      // Keep fallback behavior when backend config is unavailable.
    }
  }

  private getMapboxAccessToken(): string {
    return String(this.mapboxAccessToken || (window as any).MAPBOX_ACCESS_TOKEN || '').trim();
  }

  ngOnDestroy(): void {
    this.notificationsWsShouldRun = false;
    if (this.notificationsWsRetryTimer) window.clearTimeout(this.notificationsWsRetryTimer);
    this.notificationsWsRetryTimer = null;
    try {
      this.notificationsWs?.close();
    } catch {
      // ignore
    }
    this.notificationsWs = null;
  }

  private async apiJson<T>(method: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE', path: string, body?: unknown): Promise<T> {
    const doFetch = async () => {
      const headers: Record<string, string> = {};
      if (body != null) headers['Content-Type'] = 'application/json';
      if (this.authAccessToken) headers['Authorization'] = `Bearer ${this.authAccessToken}`;

      return fetch(`${this.apiBaseUrl}${path}`, {
        method,
        headers: Object.keys(headers).length ? headers : undefined,
        body: body == null ? undefined : JSON.stringify(body)
      });
    };

    let res = await doFetch();

    // Auto-refresh on 401 for authenticated sessions (except auth endpoints).
    if (
      res.status === 401 &&
      this.isAuthenticated &&
      !path.startsWith('/auth/login') &&
      !path.startsWith('/auth/refresh') &&
      !path.startsWith('/auth/logout')
    ) {
      const refreshed = await this.refreshAuthToken();
      if (refreshed) {
        res = await doFetch();
      } else {
        this.logoutFromSystem();
      }
    }

    const text = await res.text();
    let data: any = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = null;
    }

    if (!res.ok) {
      // Prefer backend detail for easier debugging in local.
      const msg = data?.detail || data?.message || data?.error || text || `HTTP ${res.status}`;
      const err = new Error(String(msg)) as Error & { apiCode?: string; httpStatus?: number };
      err.apiCode = typeof data?.error === 'string' ? data.error : undefined;
      err.httpStatus = Number(res.status || 0) || undefined;
      throw err;
    }

    return data as T;
  }

  private formatMoney(value: number): string {
    const safe = Number.isFinite(value) ? Math.max(0, value) : 0;
    return `$${Math.round(safe).toLocaleString('en-US')}`;
  }

  private normalizeIsoDate(iso: string | null | undefined): string {
    const raw = String(iso || '').trim();
    if (!raw) return '';
    const dt = new Date(raw);
    if (Number.isNaN(dt.getTime())) return '';
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  private weekdayLabel(date: Date): string {
    const raw = new Intl.DateTimeFormat('es-PE', { weekday: 'short' }).format(date).replace('.', '');
    const v = raw.trim();
    return v ? v.charAt(0).toUpperCase() + v.slice(1) : '';
  }

  private refreshDashboardMetrics(): void {
    const today = this.getTodayDate();

    this.todayOrders = this.orders.filter((o) => String(o.deliveryDate || '') === today).length;
    this.pendingOrders = this.orders.filter((o) => o.status === 'Pendiente').length;
    this.completedOrders = this.orders.filter((o) => o.status === 'Entregado').length;

    const salesToday = this.billingDocuments
      .filter((d) => this.normalizeIsoDate(d.createdAt) === today)
      .reduce((sum, d) => sum + Number(d.total || 0), 0);
    this.dailySales = this.formatMoney(salesToday);

    const todayDt = new Date();
    todayDt.setHours(0, 0, 0, 0);

    const dayDefs: Array<{ key: string; day: string }> = [];
    for (let i = 6; i >= 0; i -= 1) {
      const d = new Date(todayDt);
      d.setDate(todayDt.getDate() - i);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      dayDefs.push({ key, day: this.weekdayLabel(d) });
    }

    const salesByDate = new Map<string, number>();
    for (const doc of this.billingDocuments) {
      const key = this.normalizeIsoDate(doc.createdAt);
      if (!key) continue;
      salesByDate.set(key, (salesByDate.get(key) || 0) + Number(doc.total || 0));
    }

    this.salesByDay = dayDefs.map((d) => ({ day: d.day, value: Math.round(salesByDate.get(d.key) || 0) }));

    this.ordersByZone = this.zoneOptions.map((zone) => ({
      zone,
      value: this.orders.filter((o) => o.zone === zone).length
    }));

    this.maxSales = Math.max(1, ...this.salesByDay.map((item) => item.value));
    this.maxZones = Math.max(1, ...this.ordersByZone.map((item) => item.value));
  }

  private async loadClientsFromApi(): Promise<void> {
    try {
      const raw = await this.apiJson<any[]>('GET', '/clients');
      this.clients = (raw ?? []).map((c) => ({
        id: Number(c.id),
        documentType: c.documentType,
        documentNumber: String(c.documentNumber ?? ''),
        name: String(c.name ?? ''),
        address: String(c.address ?? ''),
        phone: String(c.phone ?? ''),
        email: String(c.email ?? ''),
        clientType: c.clientType
      })) as Client[];

      this.syncCreditNoteClientSelection();
    } catch {
      // Keep demo data if backend is down.
    }
  }

  private syncCreditNoteClientSelection(): void {
    // If a credit note is being created from an affected document,
    // select the real client from the list to avoid an empty/duplicate label.
    if (this.billingForm?.type !== 'Nota de crédito') return;
    if (this.billingAffectedBillingId == null) return;
    if (this.billingSelectedClientId != null) return;

    const doc = String(this.billingForm.clientDocument || '').trim();
    if (!doc) return;
    const c = (this.clients || []).find((x) => String(x.documentNumber || '').trim() === doc);
    if (c) this.billingSelectedClientId = c.id;
  }

  getBillingClientOptions(): Client[] {
    const list = this.clients || [];
    // If we're showing a fallback client label (selectedClientId == null),
    // hide the matching client to avoid duplicates in the dropdown.
    if (this.billingForm?.type === 'Nota de crédito' && this.billingSelectedClientId == null) {
      const doc = String(this.billingForm.clientDocument || '').trim();
      if (doc) return list.filter((c) => String(c.documentNumber || '').trim() !== doc);
    }
    return list;
  }

  private async loadVehiclesFromApi(): Promise<void> {
    try {
      const raw = await this.apiJson<any[]>('GET', '/vehicles');
      this.vehicles = (raw ?? []).map((v) => ({
        id: Number(v.id),
        plate: String(v.plate ?? ''),
        capacity: Number(v.capacity),
        status: v.status,
        soat: String(v.soat ?? ''),
        technicalReview: String(v.technicalReview ?? '')
      })) as Vehicle[];
      this.vehicleOptions = this.vehicles.map((v) => v.plate);
    } catch {
      // Keep demo data.
    }
  }

  private async loadDriversFromApi(): Promise<void> {
    try {
      const raw = await this.apiJson<any[]>('GET', '/drivers');
      this.drivers = (raw ?? []).map((d) => ({
        id: Number(d.id),
        name: String(d.name ?? ''),
        dni: String(d.dni ?? ''),
        license: String(d.license ?? ''),
        phone: String(d.phone ?? ''),
        status: d.status
      })) as Driver[];
      this.driverOptions = this.drivers.map((d) => d.name);
    } catch {
      // Keep demo data.
    }
  }

  private async loadOrdersFromApi(filters?: { date?: string; month?: string }): Promise<void> {
    try {
      const qs = new URLSearchParams();
      if (filters?.date) qs.set('date', filters.date);
      if (filters?.month) qs.set('month', filters.month);
      const path = `/orders${qs.toString() ? `?${qs.toString()}` : ''}`;

      const list = await this.apiJson<Array<any>>('GET', path);
      this.orders = list.map((o) => ({
        id: Number(o.id),
        clientId: o.clientId == null ? null : Number(o.clientId),
        clientName: String(o.clientName ?? ''),
        clientDocument: o.clientDocument ? String(o.clientDocument) : undefined,
        zone: o.zone as ZoneType,
        deliveryAddress: String(o.deliveryAddress ?? ''),
        serviceType: o.serviceType as ServiceType,
        quantity: o.quantity == null ? null : Number(o.quantity),
        price: o.price == null ? null : Number(o.price),
        deliveryDate: String(o.deliveryDate ?? ''),
        status: o.status as OrderStatus,
        vehicle: String(o.vehicle ?? ''),
        driver: String(o.driver ?? ''),
        history: []
      }));
      this.refreshDashboardMetrics();
      this.syncDelayedOrderNotifications();
      this.refreshRouteData();
    } catch {
      // Keep demo data.
    }
  }

  private routeOrdersLoading = false;
  private routeOrdersLoadedMonth = '';

  private async loadOrdersForRoutesView(): Promise<void> {
    const monthKey = String(this.routeDate || '').slice(0, 7); // YYYY-MM
    if (monthKey.length !== 7) return;

    // Avoid reloading same month repeatedly (unless forced).
    if (this.routeOrdersLoadedMonth === monthKey && this.orders.length) return;

    this.routeOrdersLoading = true;
    try {
      await this.loadOrdersFromApi({ month: monthKey });
      this.routeOrdersLoadedMonth = monthKey;
    } finally {
      this.routeOrdersLoading = false;
    }
  }

  private async loadBillingFromApi(): Promise<void> {
    try {
      this.billingDocuments = await this.apiJson<BillingDocument[]>('GET', '/billing-documents');
      this.refreshDashboardMetrics();
    } catch (e) {
      this.billingDocuments = [];
      this.refreshDashboardMetrics();
      this.pushToast('error', String((e as any)?.message || 'No se pudo cargar documentos de facturación (backend).'), 5000);
    }
  }

  private async loadSecurityFromApi(): Promise<void> {
    try {
      const raw = await this.apiJson<any[]>('GET', '/users');
      this.users = (raw ?? []).map((u) => ({
        id: Number(u.id),
        username: String(u.username ?? ''),
        email: String(u.email ?? ''),
        role: u.role,
        status: u.status,
        isTemporary: Boolean(u.isTemporary),
        expiresAt: u.expiresAt == null || String(u.expiresAt ?? '').trim() === '' ? null : String(u.expiresAt),
        createdAt: String(u.createdAt ?? ''),
        lastLogin: String(u.lastLogin ?? '')
      })) as User[];
    } catch {
      // Keep demo data.
    }

    try {
      const raw = await this.apiJson<any[]>('GET', '/roles');
      this.roles = (raw ?? []).map((r) => ({
        id: Number(r.id),
        name: r.name,
        description: String(r.description ?? ''),
        permissions: Array.isArray(r.permissions) ? r.permissions.map((x: any) => String(x)) : []
      })) as Role[];
    } catch {
      // Keep demo data.
    }

    try {
      const raw = await this.apiJson<any[]>('GET', '/permissions');
      this.permissions = (raw ?? []).map((p) => ({
        id: Number(p.id),
        module: String(p.module ?? ''),
        description: String(p.description ?? ''),
        type: p.type
      })) as Permission[];
    } catch {
      // Keep demo data.
    }
  }
  view: View = 'dashboard';
  private lastView: View = this.view;
  sidebarOpen = true;

  menuTooltipVisible = false;
  menuTooltipText = '';
  menuTooltipX = 0;
  menuTooltipY = 0;
  private menuTooltipAnchor: HTMLElement | null = null;

  menuPopoverVisible = false;
  menuPopoverTitle = '';
  menuPopoverItems: Array<{ label: string; action: () => void }> = [];
  menuPopoverX = 0;
  menuPopoverY = 0;
  hoveringMenuPopover = false;
  private menuPopoverAnchor: HTMLElement | null = null;
  private menuPopoverKey: string | null = null;
  private menuPopoverRaf: number | null = null;

  dashboardNotificationsOpen = false;
  private dashboardNotifSeq = 3;
  private dashboardNotifKeys = new Set<string>(['seed:1', 'seed:2', 'seed:3']);

  assignModalOpen = false;
  assignModalOrderId: number | null = null;
  assignModalVehicleId: number | null = null;
  assignModalDriverId: number | null = null;
  assignModalError = '';

  addressFixModalOpen = false;
  addressFixOrderId: number | null = null;
  addressFixMapsUrl = '';
  addressFixAddress = '';
  private addressFixOriginalAddress = '';
  addressFixResolvedLabel = '';
  addressFixError = '';
  addressFixResolving = false;
  addressFixSaving = false;

  billingFolderModalOpen = false;
  billingFolderPath = 'C:\\';

  billingLookupLoading = false;
  billingLookupMessage = '';
  billingLookupError = '';
  private billingLookupDebounceTimer: number | null = null;
  private billingLookupAbort: AbortController | null = null;
  private billingLastLookupKey = '';
  private billingLastAutoFilledName = '';

  billingClientNameEditable = true;
  billingClientNameAutoFilled = false;
  private billingClientNameManualOverride = false;
  private billingPrevClientDocument = '';

  toggleSidebar(): void {
    this.sidebarOpen = !this.sidebarOpen;
    this.hideMenuTooltip();
    this.hideMenuPopover();
  }

  onSidebarScroll(): void {
    // If the sidebar scrolls while collapsed, the anchor moves under a fixed overlay.
    // Closing avoids a "detached" popover/tooltip.
    this.hideMenuTooltip();
    this.hideMenuPopover();
  }

  onMenuHeaderClick(event: MouseEvent, menu: string): void {
    if (this.sidebarOpen) {
      this.toggleMenu(menu);
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const el = event.currentTarget as HTMLElement | null;
    if (!el) return;

    // Toggle popover when clicking the same menu.
    if (this.menuPopoverVisible && this.menuPopoverKey === menu) {
      this.hideMenuPopover();
      return;
    }

    this.showMenuPopover(el, menu);
  }

  showMenuTooltip(event: MouseEvent): void {
    if (this.sidebarOpen) return;
    if (this.menuPopoverVisible) return;
    const el = event.currentTarget as HTMLElement | null;
    if (!el) return;

    const text = el.getAttribute('data-title') ?? '';
    if (!text) return;

    this.menuTooltipAnchor = el;
    this.menuTooltipText = text;
    this.menuTooltipVisible = true;
    this.positionMenuTooltip();
  }

  moveMenuTooltip(_event: MouseEvent): void {
    if (!this.menuTooltipVisible) return;
    this.positionMenuTooltip();
  }

  hideMenuTooltip(): void {
    this.menuTooltipVisible = false;
    this.menuTooltipText = '';
    this.menuTooltipAnchor = null;
  }

  onMenuPopoverItemClick(item: { label: string; action: () => void }): void {
    item.action();
    this.hideMenuPopover();
  }

  private showMenuPopover(anchor: HTMLElement, menu: string): void {
    const title = anchor.getAttribute('data-title') ?? '';
    const items = this.getMenuPopoverItems(menu);
    if (items.length === 0) return;

    this.hideMenuTooltip();

    this.menuPopoverAnchor = anchor;
    this.menuPopoverKey = menu;
    this.menuPopoverTitle = title;
    this.menuPopoverItems = items;
    this.menuPopoverVisible = true;
    this.positionMenuPopover();

    // Re-position after render so we can clamp using actual popover height.
    if (this.menuPopoverRaf !== null) cancelAnimationFrame(this.menuPopoverRaf);
    this.menuPopoverRaf = requestAnimationFrame(() => {
      this.menuPopoverRaf = null;
      this.positionMenuPopover();
    });
  }

  private hideMenuPopover(): void {
    this.menuPopoverVisible = false;
    this.menuPopoverTitle = '';
    this.menuPopoverItems = [];
    this.menuPopoverAnchor = null;
    this.menuPopoverKey = null;
    this.hoveringMenuPopover = false;
    if (this.menuPopoverRaf !== null) {
      cancelAnimationFrame(this.menuPopoverRaf);
      this.menuPopoverRaf = null;
    }
  }

  private positionMenuPopover(): void {
    const el = this.menuPopoverAnchor;
    if (!el) return;

    const rect = el.getBoundingClientRect();
    const x = rect.right + 12;
    let y = rect.top;

    const popoverEl = document.querySelector('.menu-popover-overlay') as HTMLElement | null;
    const measuredH = popoverEl?.getBoundingClientRect().height ?? 0;
    const estimatedH = 52 + this.menuPopoverItems.length * 44 + 24;
    const h = measuredH || estimatedH;

    const pad = 8;
    this.menuPopoverX = Math.min(Math.max(x, pad), window.innerWidth - pad);
    y = Math.min(Math.max(y, pad), Math.max(pad, window.innerHeight - pad - h));
    this.menuPopoverY = y;
  }

  @HostListener('window:resize')
  onWindowResize(): void {
    if (!this.menuPopoverVisible && !this.menuTooltipVisible) return;
    this.positionMenuTooltip();
    this.positionMenuPopover();
  }

  @HostListener('window:resize')
  onWindowResizeMap(): void {
    if (!this.routesMap) return;
    requestAnimationFrame(() => this.routesMap?.resize?.());
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    const target = event.target as HTMLElement | null;
    if (!target) return;

    if (this.dashboardNotificationsOpen && !target.closest('.dashboard-notifs')) {
      this.dashboardNotificationsOpen = false;
    }

    if (!this.menuPopoverVisible) return;
    if (target.closest('.menu-popover-overlay')) return;
    if (target.closest('.menu__header')) return;
    this.hideMenuPopover();
  }

  @HostListener('document:keydown.escape')
  onEsc(): void {
    this.hideMenuTooltip();
    this.hideMenuPopover();
    this.dashboardNotificationsOpen = false;
    this.closeAssignModal();
    this.closeAddressFixModal();
    this.closeBillingFolderModal();
    this.closeCompanyPasswordModal();
    this.closeUserResetPasswordModal();
  }

  toggleDashboardNotifications(event?: Event): void {
    event?.stopPropagation();
    this.dashboardNotificationsOpen = !this.dashboardNotificationsOpen;
    if (this.dashboardNotificationsOpen) void this.loadDashboardNotificationsFromApi();
  }

  get unreadDashboardNotificationsCount(): number {
    return this.dashboardNotifications.filter((n) => !n.read).length;
  }

  async markDashboardNotificationRead(id: number): Promise<void> {
    const idx = this.dashboardNotifications.findIndex((n) => n.id === id);
    if (idx < 0) return;
    if (this.dashboardNotifications[idx].read) return;
    this.dashboardNotifications[idx] = { ...this.dashboardNotifications[idx], read: true };
    this.dashboardNotifications = [...this.dashboardNotifications];
    try {
      await this.apiJson<any>('PATCH', `/notifications/${id}/read?userId=${this.notificationUserId}`);
    } catch {
      // ignore (optimistic UI)
    }
  }

  async clearDashboardNotifications(): Promise<void> {
    const prev = this.dashboardNotifications;
    this.dashboardNotifications = [];
    this.dashboardNotifKeys.clear();
    try {
      await this.apiJson<any>('DELETE', `/notifications?userId=${this.notificationUserId}&scope=all`);
      this.pushToast('success', 'Notificaciones limpiadas.', 3000, false);
    } catch {
      this.dashboardNotifications = prev;
      this.pushToast('error', 'No se pudo limpiar notificaciones.', 5000, false);
    }
  }

  addDashboardNotification(text: string, opts?: { level?: ToastKind; key?: string }): void {
    const t = String(text || '').trim();
    if (!t) return;

    const key = String(opts?.key || '').trim();
    if (key && this.dashboardNotifKeys.has(key)) return;

    this.dashboardNotifSeq += 1;
    if (key) this.dashboardNotifKeys.add(key);
    this.dashboardNotifications = [
      {
        id: this.dashboardNotifSeq,
        text: t,
        createdAt: new Date().toISOString(),
        read: false,
        level: opts?.level || 'success',
        key: key || undefined
      },
      ...this.dashboardNotifications
    ].slice(0, 100);
  }

  private upsertDashboardNotification(n: any): void {
    const id = Number(n?.id);
    if (!Number.isFinite(id)) return;

    const row: DashboardNotification = {
      id,
      text: String(n?.message || n?.text || ''),
      createdAt: String(n?.createdAt || n?.created_at || new Date().toISOString()),
      read: Boolean(n?.read ?? n?.readAt ?? n?.read_at),
      level: n?.level === 'error' ? 'error' : 'success',
      key: n?.key || n?.dedupeKey || undefined
    };

    if (row.key) this.dashboardNotifKeys.add(String(row.key));

    const idx = this.dashboardNotifications.findIndex((x) => x.id === row.id);
    if (idx >= 0) {
      this.dashboardNotifications[idx] = row;
      this.dashboardNotifications = [...this.dashboardNotifications];
      return;
    }

    this.dashboardNotifications = [row, ...this.dashboardNotifications].slice(0, 100);
  }

  private async loadDashboardNotificationsFromApi(): Promise<void> {
    try {
      const rows = await this.apiJson<any[]>('GET', `/notifications?userId=${this.notificationUserId}&limit=100`);
      this.dashboardNotifKeys.clear();
      this.dashboardNotifications = (rows || []).map((n) => {
        const key = String(n?.key || n?.dedupeKey || '').trim();
        if (key) this.dashboardNotifKeys.add(key);
        return {
          id: Number(n?.id),
          text: String(n?.message || ''),
          createdAt: String(n?.createdAt || new Date().toISOString()),
          read: Boolean(n?.read),
          level: n?.level === 'error' ? 'error' : 'success',
          key: key || undefined
        } as DashboardNotification;
      });

      const maxId = this.dashboardNotifications.reduce((m, x) => Math.max(m, Number(x.id) || 0), 0);
      if (maxId > this.dashboardNotifSeq) this.dashboardNotifSeq = maxId;
    } catch {
      // keep local list
    }
  }

  private connectNotificationsWs(): void {
    if (!this.notificationsWsShouldRun) return;
    if (!Number.isFinite(this.notificationUserId) || this.notificationUserId < 1) return;
    if (this.notificationsWs && this.notificationsWs.readyState === WebSocket.OPEN) return;

    try {
      const wsUrl = toNotificationsWsUrl(this.apiBaseUrl, this.notificationUserId);
      const ws = new WebSocket(wsUrl);
      this.notificationsWs = ws;

      ws.onmessage = (ev) => {
        try {
          const msg = JSON.parse(String(ev.data || '{}'));
          if (msg?.type === 'notification.created' && msg.notification) {
            this.upsertDashboardNotification(msg.notification);
          } else if (msg?.type === 'notification.read' && Number.isFinite(Number(msg.id))) {
            const id = Number(msg.id);
            const idx = this.dashboardNotifications.findIndex((n) => n.id === id);
            if (idx >= 0 && !this.dashboardNotifications[idx].read) {
              this.dashboardNotifications[idx] = { ...this.dashboardNotifications[idx], read: true };
              this.dashboardNotifications = [...this.dashboardNotifications];
            }
          } else if (msg?.type === 'notification.read-all') {
            this.dashboardNotifications = this.dashboardNotifications.map((n) => ({ ...n, read: true }));
          } else if (msg?.type === 'notification.cleared') {
            this.dashboardNotifications = [];
            this.dashboardNotifKeys.clear();
          }
        } catch {
          // ignore malformed event
        }
      };

      ws.onclose = () => {
        if (!this.notificationsWsShouldRun) return;
        if (this.notificationsWsRetryTimer) window.clearTimeout(this.notificationsWsRetryTimer);
        this.notificationsWsRetryTimer = window.setTimeout(() => this.connectNotificationsWs(), 2000);
      };

      ws.onerror = () => {
        try {
          ws.close();
        } catch {
          // ignore
        }
      };
    } catch {
      if (this.notificationsWsRetryTimer) window.clearTimeout(this.notificationsWsRetryTimer);
      this.notificationsWsRetryTimer = window.setTimeout(() => this.connectNotificationsWs(), 2000);
    }
  }

  private emitDashboardNotification(message: string, level: ToastKind, key?: string): void {
    const text = String(message || '').trim();
    if (!text) return;

    void this.apiJson<any>('POST', '/notifications/emit', {
      userId: this.notificationUserId,
      message: text,
      level,
      source: 'frontend',
      dedupeKey: key || ''
    }).catch(() => {
      // Fallback local when backend/ws is down.
      this.addDashboardNotification(text, { level, key });
    });
  }

  formatDashboardNotificationTime(iso: string): string {
    const d = new Date(String(iso || ''));
    if (Number.isNaN(d.getTime())) return '';
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  private syncDelayedOrderNotifications(): void {
    const today = this.getTodayDate();
    for (const o of this.orders) {
      if (!o?.id) continue;
      if (!o.deliveryDate) continue;
      if (o.status === 'Entregado') continue;
      if (String(o.deliveryDate) >= today) continue;

      const key = `order-delayed:${o.id}:${o.deliveryDate}`;
      this.emitDashboardNotification(`Pedido #${o.id} retrasado (entrega ${o.deliveryDate})`, 'error', key);
    }
  }

  private getMenuPopoverItems(menu: string): Array<{ label: string; action: () => void }> {
    switch (menu) {
      case 'dashboard':
        return [{ label: 'Vista general', action: () => this.setView('dashboard') }];
      case 'clients':
        return [{ label: 'Gestion', action: () => this.openClients() }];
      case 'orders':
        return [
          { label: 'Lista', action: () => this.openOrdersList() },
          { label: 'Nuevo', action: () => this.openNewOrder() },
          { label: 'Historial', action: () => this.openOrderHistory() }
        ];
      case 'routes':
        return [
          { label: 'Rutas diarias', action: () => this.openRoutesDaily() },
          { label: 'Asignación', action: () => this.openRoutesAssignment() }
        ];
      case 'vehicles':
        return [
          { label: 'Lista', action: () => this.openVehiclesList() },
          { label: 'Nuevo', action: () => this.openNewVehicle() }
        ];
      case 'drivers':
        return [
          { label: 'Lista', action: () => this.openDriversList() },
          { label: 'Nuevo', action: () => this.openNewDriver() }
        ];
      case 'billing':
        return [{ label: 'Gestion', action: () => this.openBilling() }];
      case 'reports':
        return [{ label: 'Gestion', action: () => this.openReports(this.reportsTab || 'sales') }];
      case 'config':
        return [
          { label: 'Empresa', action: () => this.openConfigCompany() },
          { label: 'IGV', action: () => this.openConfigIgv() },
          { label: 'Certificado', action: () => this.openConfigCertificate() },
          { label: 'Series', action: () => this.openConfigSeries() },
        ];
      case 'users':
        return [
          { label: 'Usuarios', action: () => this.openUsersList() },
          { label: 'Nuevo usuario', action: () => this.openNewUser() },
          { label: 'Roles', action: () => this.openRolesList() },
          { label: 'Permisos', action: () => this.openPermissionsList() },
          { label: 'Auditoria', action: () => this.openSecurityAudit() }
        ];
      default:
        return [];
    }
  }

  openBilling(): void {
    // Single entry point from the menu.
    // Also refresh the certificate path used for signing.
    void this.loadCertificateFromApi();
    void this.loadSunatConfigFromApi();
    void this.loadClientsFromApi();
    this.openBillingView('billing-boletas', 'Boleta');
  }

  onBillingTypeChange(): void {
    const type = this.billingForm.type;
    if (type === 'Boleta') this.openBillingView('billing-boletas', 'Boleta');
    else if (type === 'Factura') this.openBillingView('billing-facturas', 'Factura');
    else if (type === 'Nota de crédito') this.openBillingView('billing-credit-notes', 'Nota de crédito');
    else this.openBillingView('billing-boletas', 'Nota de venta');
  }

  openBillingPrint(doc: BillingDocument): void {
    void this.openBillingPdfModal(doc);
  }

  async openBillingFolder(doc: BillingDocument): Promise<void> {
    try {
      const data = await this.apiJson<{ ok: boolean; path: string }>('POST', `/billing-documents/${doc.id}/open-folder`, {});
      this.billingFolderPath = String(data?.path || '');
      this.billingFolderModalOpen = true;
      this.pushToast('success', 'Explorador abierto.', 5000);
    } catch (e) {
      this.pushToast('error', String((e as any)?.message || 'No se pudo abrir la carpeta (backend).'), 5000);
    }
  }

  closeBillingFolderModal(): void {
    this.billingFolderModalOpen = false;
  }

  copyBillingFolderPath(): void {
    try {
      navigator.clipboard?.writeText(this.billingFolderPath);
    } catch {
      // Ignore clipboard failures.
    }
  }

  billingPdfModalOpen = false;
  billingPdfTitle = '';
  billingPdfUrl: SafeResourceUrl | null = null;
  private billingPdfBlobUrl: string | null = null;
  billingPdfLoading = false;
  billingPdfError = '';

  private async openBillingPdfModal(doc: BillingDocument): Promise<void> {
    this.billingPdfTitle = `${doc.type} - ${doc.clientName}`;
    this.billingPdfError = '';
    this.billingPdfLoading = true;
    this.billingPdfModalOpen = true;

    // Revoke previous blob url.
    if (this.billingPdfBlobUrl) {
      URL.revokeObjectURL(this.billingPdfBlobUrl);
      this.billingPdfBlobUrl = null;
    }
    this.billingPdfUrl = null;

    try {
      const resp = await fetch(`${this.apiBaseUrl}/billing-documents/${doc.id}/pdf`);
      if (!resp.ok) throw new Error('PDF no disponible.');
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      this.billingPdfBlobUrl = url;
      this.billingPdfUrl = this.sanitizer.bypassSecurityTrustResourceUrl(url);
    } catch (e) {
      this.billingPdfError = String((e as any)?.message || 'No se pudo cargar el PDF.');
    } finally {
      this.billingPdfLoading = false;
    }
  }

  closeBillingPdfModal(): void {
    this.billingPdfModalOpen = false;
    this.billingPdfTitle = '';
    this.billingPdfError = '';
    this.billingPdfLoading = false;
    this.billingPdfUrl = null;
    if (this.billingPdfBlobUrl) {
      URL.revokeObjectURL(this.billingPdfBlobUrl);
      this.billingPdfBlobUrl = null;
    }
  }

  printBillingPdf(): void {
    // Best-effort: printing is handled by the built-in PDF viewer.
    try {
      const iframe = document.getElementById('billingPdfFrame') as HTMLIFrameElement | null;
      iframe?.contentWindow?.focus();
      iframe?.contentWindow?.print();
    } catch {
      // ignore
    }
  }

  private printBillingDocument(doc: BillingDocument): void {
    const w = window.open('', '_blank', 'noopener,noreferrer,width=920,height=720');
    if (!w) return;

    const safe = (value: unknown) =>
      String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');

    const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>${safe(doc.type)} #${safe(doc.id)}</title>
  <style>
    :root { color-scheme: light; }
    body { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 28px; color: #0b1220; }
    .top { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; }
    h1 { margin: 0; font-size: 22px; }
    .pill { border: 1px solid #d7dee8; padding: 6px 10px; border-radius: 999px; font-size: 12px; white-space: nowrap; }
    .grid { margin-top: 18px; display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .card { border: 1px solid #e5eaf2; border-radius: 14px; padding: 14px; }
    .k { font-size: 12px; color: #52627a; margin-bottom: 6px; }
    .v { font-size: 14px; font-weight: 600; }
    .total { margin-top: 18px; display: flex; justify-content: flex-end; }
    .total .card { width: 260px; }
    .row { display:flex; justify-content: space-between; gap: 12px; margin-top: 8px; }
    .row strong { font-size: 14px; }
    .muted { color: #52627a; font-size: 12px; }
    @media print { body { margin: 0; } }
  </style>
</head>
<body>
  <div class="top">
    <div>
      <div class="muted">Facturación electrónica</div>
      <h1>${safe(doc.type)} · Documento #${safe(doc.id)}</h1>
      <div class="muted">Generado desde el sistema</div>
    </div>
    <div class="pill">SUNAT: ${safe(doc.sunatStatus)}</div>
  </div>

  <div class="grid">
    <div class="card">
      <div class="k">Cliente</div>
      <div class="v">${safe(doc.clientName)}</div>
      <div class="muted">${safe(doc.clientDocument)}</div>
    </div>
    <div class="card">
      <div class="k">Detalle</div>
      <div class="v">${safe(doc.detail)}</div>
      <div class="muted">Canal: ${safe(doc.channel)}</div>
    </div>
  </div>

  <div class="total">
    <div class="card">
      <div class="row"><span class="muted">Subtotal</span><strong>S/ ${safe(doc.subtotal)}</strong></div>
      <div class="row"><span class="muted">IGV</span><strong>S/ ${safe(doc.igv)}</strong></div>
      <div class="row"><span class="muted">Total</span><strong>S/ ${safe(doc.total)}</strong></div>
    </div>
  </div>

  <script>
    window.onload = () => { window.focus(); window.print(); };
  </script>
</body>
</html>`;

    w.document.open();
    w.document.write(html);
    w.document.close();
  }

  openClients(): void {
    this.setView('clients');
    void this.loadClientsFromApi();
    // When entering via menu, always start at the top of the view (ERP feel).
    requestAnimationFrame(() => {
      const content = document.querySelector('section.content') as HTMLElement | null;
      content?.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }

  openClientsList(): void {
    this.setView('clients');
    void this.loadClientsFromApi();
    this.scrollToContentId('clients-list');
  }

  openClientsNew(): void {
    this.setView('clients');
    void this.loadClientsFromApi();
    this.newClient();
    this.scrollToContentId('clients-form');
  }

  private scrollToContentId(id: string): void {
    // Defer so the section renders when switching views.
    requestAnimationFrame(() => {
      const el = document.getElementById(id);
      if (!el) return;
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  private positionMenuTooltip(): void {
    const el = this.menuTooltipAnchor;
    if (!el) return;

    const rect = el.getBoundingClientRect();
    const x = rect.right + 12;
    const y = rect.top + rect.height / 2;

    // Keep the tooltip within the viewport.
    const pad = 8;
    this.menuTooltipX = Math.min(Math.max(x, pad), window.innerWidth - pad);
    this.menuTooltipY = Math.min(Math.max(y, pad), window.innerHeight - pad);
  }

  expandedMenus: { [key: string]: boolean } = {
    dashboard: false,
    clients: false,
    orders: false,
    routes: false,
    vehicles: false,
    drivers: false,
    billing: false,
    reports: false,
    config: false,
    users: false
  };

  toggleMenu(menu: string): void {
    this.expandedMenus[menu] = !this.expandedMenus[menu];
  }

  isExpanded(menu: string): boolean {
    return this.expandedMenus[menu] ?? false;
  }

  todayOrders = 0;
  dailySales = '$0';
  pendingOrders = 0;
  completedOrders = 0;

  salesByDay = [
    { day: 'Lun', value: 0 },
    { day: 'Mar', value: 0 },
    { day: 'Mié', value: 0 },
    { day: 'Jue', value: 0 },
    { day: 'Vie', value: 0 },
    { day: 'Sáb', value: 0 },
    { day: 'Dom', value: 0 }
  ];

  ordersByZone = [
    { zone: 'Centro', value: 0 },
    { zone: 'Norte', value: 0 },
    { zone: 'Sur', value: 0 },
    { zone: 'Este', value: 0 },
    { zone: 'Oeste', value: 0 }
  ];

  dashboardNotifications: DashboardNotification[] = [
    { id: 1, text: 'Pedido #1042 - 18 min retraso', createdAt: '2026-04-24T09:18:00', read: false, level: 'error', key: 'seed:1' },
    { id: 2, text: 'Pedido #1048 - 12 min retraso', createdAt: '2026-04-24T09:24:00', read: false, level: 'error', key: 'seed:2' },
    { id: 3, text: 'Pedido #1051 - 9 min retraso', createdAt: '2026-04-24T09:31:00', read: false, level: 'error', key: 'seed:3' }
  ];

  zoneOptions: ZoneType[] = ['Centro', 'Norte', 'Sur', 'Este', 'Oeste'];

  maxSales = 1;
  maxZones = 1;

  clients: Client[] = [
    {
      id: 1,
      documentType: 'RUC',
      documentNumber: '20123456789',
      name: 'Distribuidora Andes SAC',
      address: 'Av. Grau 123, Lima',
      phone: '987654321',
      email: 'contacto@andes.com',
      clientType: 'empresa'
    },
    {
      id: 2,
      documentType: 'DNI',
      documentNumber: '45678912',
      name: 'María Torres',
      address: 'Jr. Lima 456, Lima',
      phone: '912345678',
      email: 'maria@gmail.com',
      clientType: 'persona'
    }
  ];

  clientForm: Client = this.emptyClient();
  editingId: number | null = null;
  validationMessage = '';
  validationError = '';
  lookupMessage = '';
  lookupError = '';
  lookupLoading = false;

  clientNameEditable = true;
  clientNameAutoFilled = false;
  private clientNameManualOverride = false;
  clientAddressEditable = true;
  clientAddressAutoFilled = false;
  private clientAddressManualOverride = false;
  private clientPrevDocumentNumber = '';

  private clientMessageTimer: number | null = null;
  private lookupMessageTimer: number | null = null;

  private lookupDebounceTimer: number | null = null;
  private lookupAbort: AbortController | null = null;
  private lastLookupKey = '';
  private lastAutoFilledName = '';
  private lastAutoFilledAddress = '';

  orderStatusOptions: OrderStatus[] = ['Pendiente', 'En ruta', 'Entregado'];
  serviceTypes: ServiceType[] = ['Agua potable', 'Piscina', 'Obra'];
  vehicleOptions = ['Camión A-01', 'Camión B-02', 'Camión C-03'];
  driverOptions = ['Carlos Pérez', 'Luis Rojas', 'Ana Díaz'];

  orders: Order[] = [
    {
      id: 1,
      clientId: 1,
      clientName: 'Distribuidora Andes SAC',
      zone: 'Centro',
      deliveryAddress: 'Av. Grau 123, Lima',
      serviceType: 'Agua potable',
      quantity: 8,
      price: 240,
      deliveryDate: '2026-04-16',
      status: 'Pendiente',
      vehicle: 'Camión A-01',
      driver: 'Carlos Pérez',
      history: ['Pedido creado']
    },
    {
      id: 2,
      clientId: 2,
      clientName: 'María Torres',
      zone: 'Sur',
      deliveryAddress: 'Jr. Lima 456, Lima',
      serviceType: 'Piscina',
      quantity: 12,
      price: 540,
      deliveryDate: '2026-04-17',
      status: 'En ruta',
      vehicle: 'Camión B-02',
      driver: 'Luis Rojas',
      history: ['Pedido creado', 'Asignado a vehículo y chofer']
    }
  ];

  orderForm: Order = this.emptyOrder();
  editingOrderId: number | null = null;
  orderMessage = '';
  orderError = '';
  orderHistory: Array<{ id: number; message: string }> = this.orders.flatMap((order) =>
    order.history.map((message) => ({ id: order.id, message: `Pedido #${order.id}: ${message}` }))
  );

  orderHistoryExpanded: { [orderId: number]: boolean } = {};

  toggleOrderHistory(orderId: number): void {
    this.orderHistoryExpanded[orderId] = !this.orderHistoryExpanded[orderId];
  }

  isOrderHistoryExpanded(orderId: number): boolean {
    return !!this.orderHistoryExpanded[orderId];
  }

  lastOrderHistoryMessage(order: Order): string {
    return order.history[order.history.length - 1] ?? '';
  }

  routeDate = '';
  routeGroupByZone: Array<{ zone: ZoneType; count: number }> = [];
  routeAssignments: Array<{ orderId: number; vehicle: string; driver: string }> = [];
  routeMapEnabled = false;
  routeMapError = '';
  private routesMap: any = null;
  private routesMapMarkers: any[] = [];
  private routesMapMarkerByOrderId = new Map<number, any>();
  private mapboxLoading: Promise<void> | null = null;
  private mapboxGeocodeCache = new Map<
    string,
    { lat: number; lng: number; resolvedLabel: string; precise: boolean }
  >();
  routeAssignmentOrderId: number | null = null;
  routeAssignmentVehicle = '';
  routeAssignmentDriver = '';

  vehicleStatusOptions: VehicleStatus[] = ['Disponible', 'En ruta', 'Mantenimiento'];
  vehicles: Vehicle[] = [
    {
      id: 1,
      plate: 'ABC-123',
      capacity: 10,
      status: 'Disponible',
      soat: '2026-12-10',
      technicalReview: '2026-11-05'
    },
    {
      id: 2,
      plate: 'XYZ-456',
      capacity: 12,
      status: 'En ruta',
      soat: '2026-10-18',
      technicalReview: '2026-09-30'
    }
  ];

  vehicleForm: Vehicle = this.emptyVehicle();
  editingVehicleId: number | null = null;
  vehicleMessage = '';
  vehicleError = '';

  driverStatusOptions: DriverStatus[] = ['Activo', 'Inactivo', 'Suspendido'];
  drivers: Driver[] = [
    { id: 1, name: 'Carlos Pérez', dni: '71234567', license: 'AII-b', phone: '987654321', status: 'Activo' },
    { id: 2, name: 'Luis Rojas', dni: '74561238', license: 'AIII-c', phone: '912345678', status: 'Activo' }
  ];

  driverForm: Driver = this.emptyDriver();
  editingDriverId: number | null = null;
  driverMessage = '';
  driverError = '';

  driverLookupLoading = false;
  driverLookupMessage = '';
  driverLookupError = '';
  driverNameEditable = false;
  driverNameAutoFilled = false;
  private driverNameManualOverride = false;
  private driverPrevDni = '';
  private driverDniFailCount = 0;
  private driverLastLookupKey = '';
  private driverLookupDebounceTimer: number | null = null;
  private driverLookupAbort: AbortController | null = null;

  billingTypeOptions: BillingDocType[] = ['Boleta', 'Factura', 'Nota de crédito', 'Nota de venta'];
  seriesDocTypeOptions: BillingDocType[] = ['Boleta', 'Factura', 'Nota de crédito', 'Nota de venta'];
  sunatChannels: SunatChannel[] = ['API directa SUNAT', 'Nubefact', 'Facturador SUNAT'];
  billingForm: BillingDocument = this.emptyBillingDocument();
  editingBillingId: number | null = null;
  billingMessage = '';
  billingError = '';
  // Nota de venta: solo se envia a SUNAT si el usuario lo decide.
  billingValidateSunat = false;
  billingSelectedClientId: number | null = null;
  private billingSelectedClientDocument = '';
  billingSelectedOrderId: number | null = null;

  // Nota de crédito (SUNAT): requiere documento afectado y motivo.
  billingAffectedDocType: 'Factura' | 'Boleta' = 'Factura';
  billingAffectedDocId = '';
  billingCreditNoteReasonCode = '01';
  billingCreditNoteReason = '';
  billingAffectedBillingId: number | null = null;

  // Nota de crédito: campos adicionales según código.
  billingCreditNoteNewClientDocument = '';
  billingCreditNoteNewClientName = '';
  billingCreditNoteCorrectedDetail = '';

  creditNoteReasonOptions: Array<{ code: string; label: string }> = [
    { code: '01', label: 'Anulación de la operación' },
    { code: '02', label: 'Anulación por error en el RUC' },
    { code: '03', label: 'Corrección por error en la descripción' },
    { code: '04', label: 'Descuento global' },
    { code: '05', label: 'Descuento por ítem' },
    { code: '06', label: 'Devolución total' },
    { code: '07', label: 'Devolución por ítem' },
    { code: '08', label: 'Bonificación' },
    { code: '09', label: 'Disminución en el valor' },
    { code: '10', label: 'Otros conceptos' }
  ];

  onCreditNoteReasonCodeChange(): void {
    const code = String(this.billingCreditNoteReasonCode || '').trim();

    // Clear fields not applicable to the selected reason.
    if (code !== '02') {
      this.billingCreditNoteNewClientDocument = '';
      this.billingCreditNoteNewClientName = '';
    }
    if (code !== '03') {
      this.billingCreditNoteCorrectedDetail = '';
    }
  }

  isCreditNoteReason(code: string): boolean {
    return String(this.billingCreditNoteReasonCode || '').trim() === code;
  }

  getAffectableBillingDocs(): BillingDocument[] {
    // Only documents accepted by SUNAT can be referenced.
    return (this.billingDocuments || []).filter(
      (d) => (d.type === 'Boleta' || d.type === 'Factura') && d.sunatStatus === 'Aceptado' && !!d.series && !!d.correlative
    );
  }

  onBillingAffectedDocSelected(): void {
    const id = this.billingAffectedBillingId;
    if (id == null) return;
    const d = (this.billingDocuments || []).find((x) => x.id === id);
    if (!d) return;

    this.billingAffectedDocType = d.type === 'Boleta' ? 'Boleta' : 'Factura';
    this.billingAffectedDocId = `${String(d.series || '').trim()}-${String(Number(d.correlative) || 0).padStart(4, '0')}`;

    // For Nota de crédito we default to full amount of affected document.
    this.billingForm.subtotal = Number(d.subtotal || 0);

    // Helpful defaults: keep client aligned with affected doc.
    this.billingForm.clientDocument = String(d.clientDocument || '');
    this.billingForm.clientName = String(d.clientName || '');

    // Prefer selecting the actual client if present in the list; otherwise keep null
    // so the select shows the fallback label (name + doc) instead of going blank.
    const doc = String(this.billingForm.clientDocument || '').trim();
    const c = (this.clients || []).find((x) => String(x.documentNumber || '').trim() === doc);
    this.billingSelectedClientId = c ? c.id : null;

    this.billingForm.detail = this.billingForm.detail || `NC de ${this.billingAffectedDocId}`;
  }

  private round2(n: number): number {
    const v = Number(n);
    if (!Number.isFinite(v)) return 0;
    return Math.round(v * 100) / 100;
  }

  private getBillingIgvRate(): number {
    const r = Number(this.igvRate);
    return Number.isFinite(r) && r > 0 ? r : 18;
  }

  onBillingTotalChange(): void {
    const total = Number(this.billingForm.total);
    const rate = this.getBillingIgvRate();

    if (!Number.isFinite(total) || total < 0) {
      this.billingForm.total = 0;
      this.billingForm.subtotal = 0;
      this.billingForm.igv = 0;
      return;
    }

    const subtotal = total / (1 + rate / 100);
    const igv = total - subtotal;
    this.billingForm.total = this.round2(total);
    this.billingForm.subtotal = this.round2(subtotal);
    this.billingForm.igv = this.round2(igv);
  }

  getBillingOrderOptions(): Order[] {
    const clientId = this.billingSelectedClientId;
    const list = this.orders || [];
    const filtered = clientId == null ? list : list.filter((o) => Number(o.clientId) === Number(clientId));
    // Most recent first.
    return [...filtered].sort((a, b) => Number(b.id) - Number(a.id));
  }

  onBillingOrderSelected(): void {
    const id = this.billingSelectedOrderId;
    if (id == null) return;
    const o = (this.orders || []).find((x) => x.id === id);
    if (!o) return;

    this.billingSelectedClientId = o.clientId ?? this.billingSelectedClientId;
    this.billingForm.clientDocument = String(o.clientDocument || this.billingForm.clientDocument || '');
    this.billingForm.clientName = String(o.clientName || this.billingForm.clientName || '');

    const qty = o.quantity == null ? '' : `${Number(o.quantity)} m3`;
    this.billingForm.detail = `${o.serviceType}${qty ? ' - ' + qty : ''}`;

    const total = this.round2((Number(o.price) || 0) * (Number(o.quantity) || 0));
    if (total > 0) {
      this.billingForm.total = total;
      this.onBillingTotalChange();
    }
  }

  billingClientExistsInOptions(): boolean {
    const id = this.billingSelectedClientId;
    if (id == null) return true;
    return (this.clients || []).some((c) => Number(c.id) === Number(id));
  }

  billingClientFallbackLabel(): string {
    const name = String(this.billingForm.clientName || '').trim();
    const doc = String(this.billingForm.clientDocument || '').trim();
    if (name && doc) return `${name} (${doc})`;
    return name || doc || 'Cliente';
  }

  sunatEnvironment: 'Prueba' | 'Producción' = 'Prueba';
  sunatEndpoint = 'https://e-beta.sunat.gob.pe/ol-ti-itcpfegem-beta/billService';
  sunatToken = '';
  sunatConnectionMessage = '';
  sunatConnectionError = '';

  sunatSolUser = '';
  sunatSolPass = '';
  sunatSolMessage = '';
  sunatSolError = '';
  private sunatSolDebounceTimer: number | null = null;
  sunatConfigMessage = '';
  sunatConfigError = '';
  private sunatConfigDebounceTimer: number | null = null;
  private sunatConfigLoaded = false;

  reportDateFrom = '';
  reportDateTo = '';
  reportClientId: number | null = null;
  reportDriverId: number | null = null;
  reportsTab: ReportsTab = 'sales';
  reportsLoading = false;
  reportsError = '';
  reportExportFormat: 'excel' | 'pdf' = 'excel';
  reportExportMessage = '';
  reportExportError = '';

  salesReports: SalesReport[] = [];
  orderReports: OrderReport[] = [];
  driverReports: DriverReport[] = [];
  monthlyIncomes: MonthlyIncome[] = [];

  billingDocuments: BillingDocument[] = [];

  setView(view: View): void {
    // Clean up map instances when leaving routes daily.
    if (this.view === 'routes-daily' && view !== 'routes-daily') this.destroyRoutesMap();
    this.view = view;
  }

  ngDoCheck(): void {
    // Some navigation paths set `this.view = ...` directly.
    // This makes sure we still clean up the map when leaving the routes view.
    if (this.lastView === 'routes-daily' && this.view !== 'routes-daily') {
      this.destroyRoutesMap();
    }
    this.lastView = this.view;
  }

  openOrdersList(): void {
    this.view = 'orders-list';
    void this.loadClientsFromApi();
    void this.loadOrdersFromApi();
    void this.loadVehiclesFromApi();
    void this.loadDriversFromApi();
  }

  openNewOrder(): void {
    this.orderForm = this.emptyOrder();
    this.editingOrderId = null;
    this.orderMessage = '';
    this.orderError = '';
    this.view = 'orders-new';
    void this.loadClientsFromApi();
    void this.loadVehiclesFromApi();
    void this.loadDriversFromApi();
  }

  openOrderHistory(): void {
    void this.loadOrdersFromApi().then(() => void this.loadOrderHistoryFromApi());
    this.view = 'orders-history';
  }

  openRoutesDaily(): void {
    // Always re-init cleanly (prevents stale Mapbox instances).
    this.destroyRoutesMap();

    // No default date: user explicitly picks a day.
    // Reset month cache so next selection always reloads.
    this.routeOrdersLoadedMonth = '';
    this.routeMapEnabled = true;
    this.routeMapError = '';
    this.view = 'routes-daily';

    // Load orders for the selected month, then init map.
    void this.loadOrdersForRoutesView().finally(() => {
      // Map container is created by *ngIf after view switch.
      setTimeout(() => void this.initRoutesMap(), 0);
    });
  }

  private destroyRoutesMap(): void {
    this.routesMapMarkers.forEach((m) => m.remove?.());
    this.routesMapMarkers = [];
    this.routesMapMarkerByOrderId.clear();
    this.routesMap?.remove?.();
    this.routesMap = null;
    this.routeMapError = '';
    this.routeMapEnabled = false;
  }

  openRoutesAssignment(): void {
    void this.loadOrdersFromApi();
    void this.loadVehiclesFromApi();
    void this.loadDriversFromApi();
    this.refreshRouteData();
    this.view = 'routes-assignment';
  }

  openVehiclesList(): void {
    this.view = 'vehicles-list';
    void this.loadVehiclesFromApi();
  }

  openNewVehicle(): void {
    this.vehicleForm = this.emptyVehicle();
    this.editingVehicleId = null;
    this.vehicleMessage = '';
    this.vehicleError = '';
    this.view = 'vehicles-new';
  }

  openDriversList(): void {
    this.view = 'drivers-list';
    void this.loadDriversFromApi();
  }

  openNewDriver(): void {
    this.driverForm = this.emptyDriver();
    this.editingDriverId = null;
    this.driverMessage = '';
    this.driverError = '';
    this.resetDriverDniLookup();
    // Default: lock name until DNI lookup succeeds or user fails 3 times.
    this.driverNameEditable = false;
    this.driverNameAutoFilled = false;
    this.driverNameManualOverride = false;
    this.driverPrevDni = '';
    this.view = 'drivers-new';
  }

  openBillingBoletas(): void {
    this.openBillingView('billing-boletas', 'Boleta');
  }

  openBillingFacturas(): void {
    this.openBillingView('billing-facturas', 'Factura');
  }

  openBillingCreditNotes(): void {
    this.openBillingView('billing-credit-notes', 'Nota de crédito');
  }

  openBillingSunat(): void {
    this.billingMessage = '';
    this.billingError = '';
    this.view = 'billing-sunat';
    void this.loadSunatConfigFromApi();
    void this.loadBillingFromApi();
  }

  openBillingIssued(): void {
    this.view = 'billing-issued';
    // Refresh SUNAT config so the user can see if sending is enabled.
    void this.loadSunatConfigFromApi();
    void this.loadBillingFromApi();
  }

  hasBorradorBillingDocs(): boolean {
    return this.billingDocuments.some((d) => d.sunatStatus === 'Borrador');
  }

  openReportsSales(): void {
    this.openReports('sales');
  }

  openReportsOrders(): void {
    this.openReports('orders');
  }

  openReportsDrivers(): void {
    this.openReports('drivers');
  }

  openReportsIncome(): void {
    this.openReports('income');
  }

  openReportsExport(): void {
    this.openReports(this.reportsTab || 'sales');
  }

  openReports(tab: ReportsTab = 'sales'): void {
    this.view = 'reports';
    this.reportsTab = tab;
    void this.loadClientsFromApi();
    void this.loadDriversFromApi();
    // Do not auto-select date range; user chooses explicitly.
    if (this.reportDateFrom && this.reportDateTo) void this.refreshReports();
  }

  setReportsTab(tab: ReportsTab): void {
    this.reportsTab = tab;
    // Reset irrelevant filters.
    if (tab !== 'orders') this.reportClientId = null;
    if (tab !== 'drivers') this.reportDriverId = null;
    void this.refreshReports();
  }

  async refreshReports(): Promise<void> {
    this.reportsError = '';
    this.reportsLoading = true;

    try {
      if (!this.reportDateFrom || !this.reportDateTo) {
        this.reportsError = 'Selecciona Desde y Hasta para generar el reporte.';
        return;
      }

      const qs = new URLSearchParams();
      if (this.reportDateFrom) qs.set('from', this.reportDateFrom);
      if (this.reportDateTo) qs.set('to', this.reportDateTo);
      if (this.reportsTab === 'orders' && this.reportClientId != null) qs.set('clientId', String(this.reportClientId));
      if (this.reportsTab === 'drivers' && this.reportDriverId != null) qs.set('driverId', String(this.reportDriverId));

      if (this.reportsTab === 'sales') {
        const data = await this.apiJson<{ rows: SalesReport[] }>('GET', `/reports/sales?${qs.toString()}`);
        this.salesReports = data.rows || [];
      } else if (this.reportsTab === 'orders') {
        const data = await this.apiJson<{ rows: OrderReport[] }>('GET', `/reports/orders?${qs.toString()}`);
        this.orderReports = data.rows || [];
      } else if (this.reportsTab === 'drivers') {
        const data = await this.apiJson<{ rows: DriverReport[] }>('GET', `/reports/drivers?${qs.toString()}`);
        this.driverReports = data.rows || [];
      } else {
        const data = await this.apiJson<{ rows: MonthlyIncome[] }>('GET', `/reports/income?${qs.toString()}`);
        this.monthlyIncomes = (data.rows || []).map((x) => ({ month: String((x as any).month ?? ''), income: Number((x as any).income ?? 0) }));
      }
    } catch (e) {
      this.reportsError = String((e as any)?.message || 'No se pudo cargar reportes.');
    } finally {
      this.reportsLoading = false;
    }
  }

  private async downloadReportExport(format: 'csv' | 'pdf'): Promise<void> {
    const qs = new URLSearchParams();
    qs.set('report', this.reportsTab);
    qs.set('format', format);
    if (this.reportDateFrom) qs.set('from', this.reportDateFrom);
    if (this.reportDateTo) qs.set('to', this.reportDateTo);
    if (this.reportsTab === 'orders' && this.reportClientId != null) qs.set('clientId', String(this.reportClientId));
    if (this.reportsTab === 'drivers' && this.reportDriverId != null) qs.set('driverId', String(this.reportDriverId));

    const url = `${this.apiBaseUrl}/reports/export?${qs.toString()}`;
    const resp = await fetch(url);
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(text || `HTTP ${resp.status}`);
    }

    const blob = await resp.blob();
    const cd = resp.headers.get('content-disposition') || '';
    const m = cd.match(/filename="([^"]+)"/i);
    const fileName = m ? m[1] : `reporte-${this.reportsTab}.${format === 'csv' ? 'csv' : 'pdf'}`;

    const a = document.createElement('a');
    const objectUrl = URL.createObjectURL(blob);
    a.href = objectUrl;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);
  }

  private async downloadReportBundle(): Promise<void> {
    const qs = new URLSearchParams();
    if (this.reportDateFrom) qs.set('from', this.reportDateFrom);
    if (this.reportDateTo) qs.set('to', this.reportDateTo);
    if (this.reportClientId != null) qs.set('clientId', String(this.reportClientId));
    if (this.reportDriverId != null) qs.set('driverId', String(this.reportDriverId));

    const url = `${this.apiBaseUrl}/reports/export-bundle?${qs.toString()}`;
    const resp = await fetch(url);
    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(text || `HTTP ${resp.status}`);
    }

    const blob = await resp.blob();
    const cd = resp.headers.get('content-disposition') || '';
    const m = cd.match(/filename="([^"]+)"/i);
    const fileName = m ? m[1] : `reportes.zip`;

    const a = document.createElement('a');
    const objectUrl = URL.createObjectURL(blob);
    a.href = objectUrl;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);
  }

  async exportAllReports(): Promise<void> {
    this.reportExportMessage = '';
    this.reportExportError = '';
    try {
      await this.downloadReportBundle();
      this.pushToast('success', 'Paquete completo exportado (ZIP: CSV + PDF + manifest).', 6000);
    } catch (e) {
      const msg = String((e as any)?.message || 'No se pudo exportar el paquete.');
      this.reportExportError = msg;
      this.pushToast('error', msg, 6000);
    }
  }

  async exportActiveReport(format: 'excel' | 'pdf'): Promise<void> {
    this.reportExportMessage = '';
    this.reportExportError = '';
    try {
      await this.downloadReportExport(format === 'excel' ? 'csv' : 'pdf');
      this.pushToast('success', format === 'excel' ? 'Reporte exportado (CSV para Excel).' : 'Reporte exportado (PDF).', 5000);
    } catch (e) {
      const msg = String((e as any)?.message || 'No se pudo exportar el reporte.');
      this.reportExportError = msg;
      this.pushToast('error', msg, 5000);
    }
  }

  generateSalesReport(): void {
    const from = this.reportDateFrom || '2026-01-01';
    const to = this.reportDateTo || '2026-12-31';

    const filtered = this.orders.filter(
      (order) => order.deliveryDate >= from && order.deliveryDate <= to
    );

    const grouped = new Map<string, { totalSales: number; orderCount: number }>();
    filtered.forEach((order) => {
      const existing = grouped.get(order.deliveryDate) || { totalSales: 0, orderCount: 0 };
      grouped.set(order.deliveryDate, {
        totalSales: existing.totalSales + (order.price ?? 0) * (order.quantity ?? 0),
        orderCount: existing.orderCount + 1
      });
    });

    this.salesReports = Array.from(grouped.entries()).map(([date, data]) => ({
      date,
      totalSales: data.totalSales,
      orderCount: data.orderCount
    }));
  }

  generateOrderReport(): void {
    if (this.reportClientId === null) {
      const grouped = new Map<
        string,
        { clientName: string; clientDocument: string; orderCount: number; totalAmount: number }
      >();

      this.orders.forEach((order) => {
        const clientId = order.clientId ?? null;
        const client = clientId ? this.clients.find((c) => c.id === clientId) : undefined;
        const clientName = client?.name || order.clientName || 'Cliente';
        const clientDocument = client?.documentNumber || '';

        // Prefer stable grouping by client id (when available).
        const key = clientId ? `id:${clientId}` : `name:${clientName}`;
        const existing = grouped.get(key) || { clientName, clientDocument, orderCount: 0, totalAmount: 0 };

        grouped.set(key, {
          clientName: existing.clientName || clientName,
          clientDocument: existing.clientDocument || clientDocument,
          orderCount: existing.orderCount + 1,
          totalAmount: existing.totalAmount + (order.price ?? 0) * (order.quantity ?? 0)
        });
      });

      this.orderReports = Array.from(grouped.values()).map((data) => ({
        clientName: data.clientName,
        clientDocument: data.clientDocument,
        orderCount: data.orderCount,
        totalAmount: data.totalAmount
      }));
    } else {
      const client = this.clients.find((c) => c.id === this.reportClientId);
      if (!client) return;

      const filtered = this.orders.filter((o) => o.clientId === this.reportClientId);
      this.orderReports = [
        {
          clientName: client.name,
          clientDocument: client.documentNumber,
          orderCount: filtered.length,
          totalAmount: filtered.reduce((sum, o) => sum + (o.price ?? 0) * (o.quantity ?? 0), 0)
        }
      ];
    }
  }

  generateDriverReport(): void {
    this.driverReports = this.drivers.map((driver) => {
      const driverOrders = this.orders.filter((o) => o.driver === driver.name);
      const completed = driverOrders.filter((o) => o.status === 'Entregado').length;
      const delivered = driverOrders.length;
      const rating = Math.round((3.5 + Math.random() * 1.5) * 10) / 10;
      const onTime = Math.round(70 + Math.random() * 30);

      return {
        driverName: driver.name,
        completedOrders: completed,
        totalDeliveries: delivered,
        rating,
        onTimePercentage: onTime
      };
    });
  }

  generateMonthlyIncome(): void {
    const months = [
      'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
      'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'
    ];

    this.monthlyIncomes = months.map((month, index) => ({
      month,
      income: Math.round(5000 + Math.random() * 15000)
    }));
  }

  exportReport(format: 'excel' | 'pdf'): void {
    this.reportExportMessage = '';
    this.reportExportError = '';

    if (format === 'excel') {
      this.setTransientNotice(
        'reportExportMessage',
        'reportExportError',
        'success',
        'Reporte exportado a Excel correctamente.',
        5000
      );
    } else {
      this.setTransientNotice(
        'reportExportMessage',
        'reportExportError',
        'success',
        'Reporte exportado a PDF correctamente.',
        5000
      );
    }
  }

  companyData: CompanyData = {
    name: '',
    ruc: '',
    address: '',
    ubigeo: '',
    phone: '',
    email: '',
    urbanizacion: '',
    distrito: '',
    provincia: '',
    departamento: ''
  };

  private companyInitialSnapshot = '';
  companyConfigured = false;
  companyEditUnlocked = false;
  private companyEditPassword = '';
  companyPasswordModalOpen = false;
  companyPasswordValue = '';
  companyPasswordError = '';

  // Series are handled by backend (Node/Express).

  igvRate = 18;
  igvMessage = '';
  igvError = '';

  certificateConfig: CertificateConfig = {
    path: '/certificados/firma.pfx',
    password: '',
    expiration: '2027-12-31'
  };

  certificateUploadLoading = false;

  // SUNAT API config is handled by backend (Node/Express).

  configMessage = '';
  configError = '';

  openConfigCompany(): void {
    this.view = 'config-company';
    this.companyEditUnlocked = false;
    this.companyEditPassword = '';
    void this.loadCompanyConfigFromApi();
  }

  private serializeCompanyData(data: CompanyData): string {
    return JSON.stringify({
      name: String(data.name || '').trim(),
      ruc: String(data.ruc || '').trim(),
      address: String(data.address || '').trim(),
      ubigeo: String(data.ubigeo || '').trim(),
      phone: String(data.phone || '').trim(),
      email: String(data.email || '').trim(),
      urbanizacion: String(data.urbanizacion || '').trim(),
      distrito: String(data.distrito || '').trim(),
      provincia: String(data.provincia || '').trim(),
      departamento: String(data.departamento || '').trim()
    });
  }

  private computeCompanyConfigured(data: CompanyData): boolean {
    const d = {
      name: String(data.name || '').trim(),
      ruc: String(data.ruc || '').trim(),
      address: String(data.address || '').trim(),
      ubigeo: String(data.ubigeo || '').trim(),
      phone: String(data.phone || '').trim(),
      email: String(data.email || '').trim(),
      urbanizacion: String(data.urbanizacion || '').trim(),
      distrito: String(data.distrito || '').trim(),
      provincia: String(data.provincia || '').trim(),
      departamento: String(data.departamento || '').trim()
    };
    return Boolean(d.name || d.ruc || d.address || d.ubigeo || d.phone || d.email || d.distrito || d.provincia || d.departamento);
  }

  canEditCompanyData(): boolean {
    return !this.companyConfigured || this.companyEditUnlocked;
  }

  async unlockCompanyData(): Promise<void> {
    this.openCompanyPasswordModal();
  }

  openCompanyPasswordModal(): void {
    this.companyPasswordModalOpen = true;
    this.companyPasswordValue = '';
    this.companyPasswordError = '';
  }

  closeCompanyPasswordModal(): void {
    this.companyPasswordModalOpen = false;
    this.companyPasswordValue = '';
    this.companyPasswordError = '';
  }

  async confirmCompanyPassword(): Promise<void> {
    const password = String(this.companyPasswordValue || '').trim();
    if (!password) {
      this.companyPasswordError = 'Ingresa password.';
      return;
    }

    this.configMessage = '';
    this.configError = '';
    try {
      await this.apiJson<any>('POST', '/config/company/unlock', { password });
      this.companyEditUnlocked = true;
      this.companyEditPassword = password;
      this.closeCompanyPasswordModal();
      this.setTransientNotice('configMessage', 'configError', 'success', 'Edición habilitada para datos de empresa.', 4000);
    } catch {
      this.companyEditUnlocked = false;
      this.companyEditPassword = '';
      this.companyPasswordError = 'Password inválido.';
    }
  }

  cancelCompanyEdit(): void {
    try {
      this.companyData = JSON.parse(this.companyInitialSnapshot || '{}') as CompanyData;
    } catch {
      // ignore
    }
    this.companyEditUnlocked = false;
    this.companyEditPassword = '';
    this.closeCompanyPasswordModal();
    this.configMessage = '';
    this.configError = '';
  }

  openConfigIgv(): void {
    this.view = 'config-igv';
    void this.loadIgvFromApi();
  }

  openConfigCertificate(): void {
    this.view = 'config-certificate';
    void this.loadCertificateFromApi();
    void this.loadSunatConfigFromApi();
  }

  private async loadCompanyConfigFromApi(): Promise<void> {
    try {
      const data = await this.apiJson<any>('GET', '/config/company');
      if (data) {
        this.companyData = {
          name: String(data.name ?? ''),
          ruc: String(data.ruc ?? ''),
          address: String(data.address ?? ''),
          ubigeo: String(data.ubigeo ?? ''),
          phone: String(data.phone ?? ''),
          email: String(data.email ?? ''),
          urbanizacion: String(data.urbanizacion ?? ''),
          distrito: String(data.distrito ?? ''),
          provincia: String(data.provincia ?? ''),
          departamento: String(data.departamento ?? '')
        };
      } else {
        this.companyData = {
          name: '',
          ruc: '',
          address: '',
          ubigeo: '',
          phone: '',
          email: '',
          urbanizacion: '',
          distrito: '',
          provincia: '',
          departamento: ''
        };
      }
      this.companyConfigured = this.computeCompanyConfigured(this.companyData);
      this.companyInitialSnapshot = this.serializeCompanyData(this.companyData);
      this.companyEditUnlocked = false;
      this.companyEditPassword = '';
    } catch {
      // Keep current state.
    }
  }

  private async loadIgvFromApi(): Promise<void> {
    try {
      const data = await this.apiJson<any>('GET', '/config/igv');
      const rate = Number(data?.igvRate);
      if (Number.isFinite(rate)) this.igvRate = rate;
    } catch {
      // Keep demo.
    }
  }

  private async loadCertificateFromApi(): Promise<void> {
    try {
      const data = await this.apiJson<any>('GET', '/config/certificate');
      if (data) {
        const expRaw = String(data.expiration ?? '').trim();
        const expMatch = expRaw.match(/\d{4}-\d{2}-\d{2}/);
        this.certificateConfig = {
          path: String(data.path ?? ''),
          // For security, do not preload passwords.
          password: '',
          expiration: expMatch ? expMatch[0] : ''
        };

        // Keep billing form in sync with the configured certificate.
        const certPath = this.certificateConfig.path.trim();
        if (certPath) {
          this.billingForm.certificatePath = certPath;
        }
      }
    } catch {
      // Keep demo.
    }
  }

  openConfigSeries(): void {
    this.view = 'config-series';
    void this.loadDocumentSeries();
  }

  seriesConfigs: DocumentSeriesConfig[] = [];

  seriesScopeOptions: Array<{
    docType: BillingDocType;
    scope: string;
    title: string;
    subtitle: string;
    placeholder: string;
  }> = [
    { docType: 'Boleta', scope: '', title: 'Boleta', subtitle: 'Serie inicia con B', placeholder: 'Ej: B001' },
    { docType: 'Factura', scope: '', title: 'Factura', subtitle: 'Serie inicia con F', placeholder: 'Ej: F001' },
    {
      docType: 'Nota de crédito',
      scope: 'Factura',
      title: 'Nota de crédito (a Factura)',
      subtitle: 'SUNAT: serie inicia con F',
      placeholder: 'Ej: FC01'
    },
    {
      docType: 'Nota de crédito',
      scope: 'Boleta',
      title: 'Nota de crédito (a Boleta)',
      subtitle: 'SUNAT: serie inicia con B',
      placeholder: 'Ej: BC01'
    },
    { docType: 'Nota de venta', scope: '', title: 'Nota de venta', subtitle: 'Uso interno', placeholder: 'Ej: NV01' }
  ];

  private seriesKey(docType: BillingDocType, scope: string): string {
    return `${docType}::${String(scope || '')}`;
  }

  seriesForm: Record<string, { series: string; nextCorrelative: number }> = {};
  seriesUnlocked: Record<string, boolean> = {};
  private seriesUnlockedPassword: Record<string, string> = {};

  seriesPasswordModalOpen = false;
  seriesPasswordTargetType: BillingDocType | null = null;
  seriesPasswordTargetScope = '';
  seriesPasswordValue = '';
  seriesPasswordMessage = '';
  seriesPasswordError = '';

  seriesMessage = '';
  seriesError = '';

  private seedSeriesFormFromConfigs(): void {
    const byKey = new Map<string, DocumentSeriesConfig>();
    this.seriesConfigs.forEach((c) => byKey.set(this.seriesKey(c.docType, String(c.scope || '')), c));

    this.seriesScopeOptions.forEach((opt) => {
      const k = this.seriesKey(opt.docType, opt.scope);
      const existing = byKey.get(k);
      const prev = this.seriesForm[k] || { series: '', nextCorrelative: 1 };
      this.seriesForm[k] = {
        series: existing?.series ?? prev.series,
        nextCorrelative: existing?.nextCorrelative ?? prev.nextCorrelative
      };
      if (this.seriesUnlocked[k] == null) this.seriesUnlocked[k] = false;
    });
  }

  isSeriesUnlocked(docType: BillingDocType, scope: string): boolean {
    return this.seriesUnlocked[this.seriesKey(docType, scope)] === true;
  }

  private lockAllSeries(): void {
    Object.keys(this.seriesUnlocked).forEach((k) => {
      this.seriesUnlocked[k] = false;
      delete this.seriesUnlockedPassword[k];
    });
  }

  openSeriesPasswordModal(type: BillingDocType, scope: string): void {
    this.seriesPasswordTargetType = type;
    this.seriesPasswordTargetScope = scope;
    this.seriesPasswordValue = '';
    this.seriesPasswordMessage = '';
    this.seriesPasswordError = '';
    this.seriesPasswordModalOpen = true;
  }

  closeSeriesPasswordModal(): void {
    this.seriesPasswordModalOpen = false;
    this.seriesPasswordTargetType = null;
    this.seriesPasswordTargetScope = '';
    this.seriesPasswordValue = '';
    this.seriesPasswordMessage = '';
    this.seriesPasswordError = '';
  }

  async confirmSeriesPassword(): Promise<void> {
    const type = this.seriesPasswordTargetType;
    const scope = String(this.seriesPasswordTargetScope || '');
    const password = this.seriesPasswordValue;
    this.seriesPasswordMessage = '';
    this.seriesPasswordError = '';

    if (!type) return;
    if (!password.trim()) {
      this.seriesPasswordError = 'Ingresa la contraseña.';
      return;
    }

    try {
      const response = await fetch(`${this.apiBaseUrl}/document-series/check-password`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password })
      });
      if (!response.ok) {
        this.seriesPasswordError = 'No se pudo validar la contraseña.';
        return;
      }

      const data = (await response.json()) as { ok?: boolean };
      if (!data.ok) {
        this.seriesPasswordError = 'Contraseña incorrecta.';
        return;
      }

      const key = this.seriesKey(type, scope);
      this.seriesUnlocked[key] = true;
      this.seriesUnlockedPassword[key] = password;
      this.seriesPasswordMessage = 'Validado correctamente.';

      setTimeout(() => this.closeSeriesPasswordModal(), 650);
    } catch {
      // Demo/offline mode
      if (password === 'AluQen+') {
        const key = this.seriesKey(type, scope);
        this.seriesUnlocked[key] = true;
        this.seriesUnlockedPassword[key] = password;
        this.seriesPasswordMessage = 'Validado correctamente.';
        setTimeout(() => this.closeSeriesPasswordModal(), 650);
        return;
      }

      this.seriesPasswordError = 'No se pudo validar la contraseña.';
    }
  }

  isSeriesConfigured(docType: BillingDocType, scope: string): boolean {
    const s = String(scope || '');
    return this.seriesConfigs.some((c) => c.docType === docType && String(c.scope || '') === s);
  }

  async loadDocumentSeries(): Promise<void> {
    this.seriesMessage = '';
    this.seriesError = '';

    try {
      const response = await fetch(`${this.apiBaseUrl}/document-series`);
      if (!response.ok) throw new Error('No se pudo cargar series desde el backend.');
      const data = (await response.json()) as { items?: DocumentSeriesConfig[] };
      this.seriesConfigs = (data.items ?? []).map((x) => ({
        docType: x.docType,
        scope: String((x as any).scope || ''),
        series: x.series,
        nextCorrelative: Number(x.nextCorrelative)
      }));
      this.seedSeriesFormFromConfigs();
      this.lockAllSeries();
      return;
    } catch {
      // Fallback local (demo) cuando no hay backend.
      const raw = localStorage.getItem('documentSeries');
      this.seriesConfigs = raw ? (JSON.parse(raw) as DocumentSeriesConfig[]) : [];
      this.seedSeriesFormFromConfigs();
      this.lockAllSeries();
      this.seriesError = 'No hay conexión con el backend. Mostrando configuración local (demo).';
    }
  }

  async saveSeries(docType: BillingDocType, scope: string): Promise<void> {
    this.seriesMessage = '';
    this.seriesError = '';

    const key = this.seriesKey(docType, scope);
    const form = this.seriesForm[key];
    const series = String(form?.series ?? '').trim();
    const nextCorrelative = Number(form?.nextCorrelative ?? 1);

    if (!series) {
      this.setTransientNotice('seriesMessage', 'seriesError', 'error', 'Ingresa la serie.', 5000);
      return;
    }
    if (!Number.isFinite(nextCorrelative) || nextCorrelative < 1) {
      this.setTransientNotice('seriesMessage', 'seriesError', 'error', 'El correlativo debe ser >= 1.', 5000);
      return;
    }

    const configured = this.isSeriesConfigured(docType, scope);
    const password = configured ? (this.seriesUnlockedPassword[key] ?? '') : null;
    if (configured && !this.isSeriesUnlocked(docType, scope)) {
      this.setTransientNotice('seriesMessage', 'seriesError', 'error', 'Para editar, primero valida la contraseña.', 5000);
      return;
    }

    try {
      const response = await fetch(`${this.apiBaseUrl}/document-series/set`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ docType, scope, series, nextCorrelative, password })
      });

      if (response.status === 401) {
        this.setTransientNotice('seriesMessage', 'seriesError', 'error', 'Password incorrecto.', 5000);
        return;
      }

      if (!response.ok) {
        const data = (await response.json().catch(() => ({}))) as { error?: string };
        this.setTransientNotice('seriesMessage', 'seriesError', 'error', data.error || 'No se pudo guardar la serie.', 5000);
        return;
      }

      await this.loadDocumentSeries();
      this.seriesUnlocked[key] = false;
      delete this.seriesUnlockedPassword[key];
      this.setTransientNotice('seriesMessage', 'seriesError', 'success', 'Serie guardada correctamente.', 5000);
    } catch {
      // Fallback local (demo)
      if (configured && password !== 'AluQen+') {
        this.setTransientNotice('seriesMessage', 'seriesError', 'error', 'Password incorrecto.', 5000);
        return;
      }

      const updated: DocumentSeriesConfig[] = this.isSeriesConfigured(docType, scope)
        ? this.seriesConfigs.map((c) =>
            c.docType === docType && String(c.scope || '') === String(scope || '')
              ? { docType, scope, series, nextCorrelative }
              : c
          )
        : [{ docType, scope, series, nextCorrelative }, ...this.seriesConfigs];

      localStorage.setItem('documentSeries', JSON.stringify(updated));
      this.seriesConfigs = updated;
      this.seedSeriesFormFromConfigs();
      this.seriesUnlocked[key] = false;
      delete this.seriesUnlockedPassword[key];
      this.setTransientNotice('seriesMessage', 'seriesError', 'success', 'Serie guardada en modo local (demo).', 5000);
    }
  }

  cancelSeriesEdit(docType: BillingDocType, scope: string): void {
    const key = this.seriesKey(docType, scope);
    this.seriesUnlocked[key] = false;
    delete this.seriesUnlockedPassword[key];
    // Revert to last saved values.
    this.seedSeriesFormFromConfigs();
  }


  async saveCompanyData(): Promise<void> {
    this.configMessage = '';
    this.configError = '';

    if (this.companyConfigured && !this.companyEditUnlocked) {
      this.setTransientNotice('configMessage', 'configError', 'error', 'Datos bloqueados. Usa Editar e ingresa password.', 5000);
      return;
    }

    try {
      await this.apiJson<any>('PUT', '/config/company', {
        ...this.companyData,
        password: this.companyEditPassword
      });
      this.setTransientNotice('configMessage', 'configError', 'success', 'Datos de empresa guardados correctamente.', 5000);
      await this.loadCompanyConfigFromApi();
    } catch (e) {
      const msg = String((e as any)?.message || 'No se pudo guardar la empresa (backend).');
      if (msg.includes('PASSWORD_REQUIRED')) {
        this.setTransientNotice('configMessage', 'configError', 'error', 'Password requerido para editar datos de empresa.', 5000);
      } else {
        this.setTransientNotice('configMessage', 'configError', 'error', msg, 5000);
      }
    }
  }


  async saveIgvRate(): Promise<void> {
    if (this.igvRate < 0 || this.igvRate > 100) {
      this.setTransientNotice('igvMessage', 'igvError', 'error', 'El IGV debe estar entre 0 y 100.', 5000);
      return;
    }
    this.igvMessage = '';
    this.igvError = '';
    try {
      await this.apiJson<any>('PUT', '/config/igv', { igvRate: this.igvRate });
      this.setTransientNotice('igvMessage', 'igvError', 'success', 'IGV guardado correctamente.', 5000);
      await this.loadIgvFromApi();
    } catch {
      this.setTransientNotice('igvMessage', 'igvError', 'error', 'No se pudo guardar el IGV (backend).', 5000);
    }
  }

  async saveCertificate(): Promise<void> {
    if (!this.certificateConfig.path) {
      this.setTransientNotice('configMessage', 'configError', 'error', 'Ingresa la ruta del certificado.', 5000);
      return;
    }
    this.configMessage = '';
    this.configError = '';
    try {
      await this.apiJson<any>('PUT', '/config/certificate', { ...this.certificateConfig });
      this.setTransientNotice(
        'configMessage',
        'configError',
        'success',
        'Certificado digital configurado correctamente.',
        5000
      );
      await this.loadCertificateFromApi();
    } catch {
      this.setTransientNotice('configMessage', 'configError', 'error', 'No se pudo guardar el certificado (backend).', 5000);
    }
  }

  pickCertificateFile(input: HTMLInputElement): void {
    if (this.certificateUploadLoading) return;
    try {
      input.value = '';
      input.click();
    } catch {
      // Ignore.
    }
  }

  async onCertificateFileSelected(event: Event): Promise<void> {
    const el = event.target as HTMLInputElement | null;
    const file = el?.files?.[0];
    if (!file) return;

    if (!/\.pfx$/i.test(file.name)) {
      this.pushToast('error', 'Solo se permite archivos .pfx', 5000);
      return;
    }

    this.certificateUploadLoading = true;
    try {
      // New file: clear previous expiration and require password re-entry.
      this.certificateConfig.expiration = '';

      const password = String(this.certificateConfig.password || '');
      if (!password) {
        this.pushToast('error', 'Ingresa la contraseña del certificado antes de cargar el .pfx.', 5000);
        return;
      }

      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('No se pudo leer el archivo.'));
        reader.onload = () => resolve(String(reader.result || ''));
        reader.readAsDataURL(file);
      });

      const comma = dataUrl.indexOf(',');
      const base64 = comma >= 0 ? dataUrl.slice(comma + 1) : '';
      if (!base64) throw new Error('Archivo invalido.');

      const out = await this.apiJson<{ ok: boolean; path: string; folder: string; expiration?: string }>(
        'POST',
        '/config/certificate/upload',
        {
          fileName: file.name,
          base64,
          password,
        }
      );

      if (out?.path) {
        this.certificateConfig.path = out.path;
        if (out.expiration) this.certificateConfig.expiration = String(out.expiration);
        this.pushToast('success', `Certificado cargado en: ${out.folder}`, 5000);
        // Clear password after use.
        this.certificateConfig.password = '';
        await this.loadCertificateFromApi();
      }
    } catch (e) {
      // Force re-entry on failures.
      this.certificateConfig.password = '';
      this.certificateConfig.expiration = '';
      this.pushToast('error', String((e as any)?.message || 'No se pudo subir el certificado.'), 5000);
    } finally {
      this.certificateUploadLoading = false;
    }
  }


  roleOptions: UserRole[] = ['Admin', 'Operador', 'Chofer'];
  userStatusOptions: UserStatus[] = ['Activo', 'Inactivo', 'Suspendido'];
  permissionTypes: PermissionType[] = ['read', 'write', 'delete', 'admin'];

  users: User[] = [
    { id: 1, username: 'admin', email: 'admin@empresa.com', role: 'Admin', status: 'Activo', createdAt: '2026-01-01', lastLogin: '2026-04-16' },
    { id: 2, username: 'operador1', email: 'operador1@empresa.com', role: 'Operador', status: 'Activo', createdAt: '2026-01-15', lastLogin: '2026-04-15' },
    { id: 3, username: 'chofer1', email: 'chofer1@empresa.com', role: 'Chofer', status: 'Activo', createdAt: '2026-02-01', lastLogin: '2026-04-14' }
  ];

  roles: Role[] = [
    {
      id: 1,
      name: 'Admin',
      description: 'Acceso completo al sistema',
      permissions: ['all']
    },
    {
      id: 2,
      name: 'Operador',
      description: 'Gestión de pedidos, clientes y facturación',
      permissions: ['clients', 'orders', 'billing']
    },
    {
      id: 3,
      name: 'Chofer',
      description: 'Solo visualización de rutas asignadas',
      permissions: ['routes']
    }
  ];

  permissions: Permission[] = [
    { id: 1, module: 'dashboard', description: 'Vista general', type: 'read' },
    { id: 2, module: 'clients', description: 'Gestión de clientes', type: 'write' },
    { id: 3, module: 'orders', description: 'Gestión de pedidos', type: 'write' },
    { id: 4, module: 'billing', description: 'Facturación electrónica', type: 'write' },
    { id: 5, module: 'reports', description: 'Reportes', type: 'read' },
    { id: 6, module: 'config', description: 'Configuración', type: 'admin' },
    { id: 7, module: 'users', description: 'Gestión de usuarios', type: 'admin' },
    { id: 8, module: 'routes', description: 'Rutas y asignaciones', type: 'read' }
  ];

  userForm: User = this.emptyUser();
  editingUserId: number | null = null;
  userMessage = '';
  userError = '';

  userResetPasswordModalOpen = false;
  userResetPasswordUserId: number | null = null;
  userResetPasswordUsername = '';
  userResetPasswordValue = '';
  userResetPasswordConfirm = '';
  userResetPasswordError = '';

  roleForm: Role = this.emptyRole();
  editingRoleId: number | null = null;
  roleMessage = '';
  roleError = '';

  auditLogs: AuthAuditLog[] = [];
  auditLoading = false;
  auditError = '';
  auditFilterEventType = '';
  auditFilterSuccess: '' | '1' | '0' = '';
  auditFilterFrom = '';
  auditFilterTo = '';

  openUsersList(): void {
    this.view = 'users-list';
    void this.loadSecurityFromApi();
  }

  openNewUser(): void {
    this.userForm = this.emptyUser();
    this.editingUserId = null;
    this.userMessage = '';
    this.userError = '';
    this.view = 'users-new';
  }

  private validateUserPasswordPolicy(password: string): string | null {
    const p = String(password || '');
    if (p.length < 8) return 'La contraseña debe tener al menos 8 caracteres.';
    if (p.length > 72) return 'La contraseña no debe superar 72 caracteres.';
    if (!/[A-Za-z]/.test(p)) return 'La contraseña debe incluir al menos una letra.';
    if (!/\d/.test(p)) return 'La contraseña debe incluir al menos un número.';
    return null;
  }

  private toIsoFromLocalDatetime(localValue: string | null | undefined): string | null {
    const raw = String(localValue || '').trim();
    if (!raw) return null;
    const dt = new Date(raw);
    if (Number.isNaN(dt.getTime())) return null;
    return dt.toISOString();
  }

  private toLocalDatetimeInput(isoValue: string | null | undefined): string {
    const raw = String(isoValue || '').trim();
    if (!raw) return '';
    const dt = new Date(raw);
    if (Number.isNaN(dt.getTime())) return '';
    const offsetMs = dt.getTimezoneOffset() * 60000;
    return new Date(dt.getTime() - offsetMs).toISOString().slice(0, 16);
  }

  userTypeLabel(user: User): string {
    return user.isTemporary ? 'Temporal' : 'Permanente';
  }

  onUserTemporaryChanged(): void {
    if (!this.userForm.isTemporary) {
      this.userForm.expiresAt = '';
      this.userForm.temporaryHours = null;
    }
  }

  formatUserExpiresAt(expiresAt: string | null | undefined): string {
    const iso = String(expiresAt || '').trim();
    if (!iso) return '-';
    const dt = new Date(iso);
    if (Number.isNaN(dt.getTime())) return '-';
    return dt.toLocaleString('es-PE');
  }

  formatAuthExpiresAt(): string {
    return this.formatUserExpiresAt(this.authExpiresAt);
  }

  openRolesList(): void {
    this.view = 'roles-list';
    void this.loadSecurityFromApi();
  }

  openNewRole(): void {
    this.roleForm = this.emptyRole();
    this.editingRoleId = null;
    this.roleMessage = '';
    this.roleError = '';
    this.view = 'roles-list';
  }

  openPermissionsList(): void {
    this.view = 'permissions-list';
    void this.loadSecurityFromApi();
  }

  openPermissionsEdit(): void {
    this.view = 'permissions-edit';
  }

  openSecurityAudit(): void {
    this.view = 'security-audit';
    void this.loadAuthAuditLogs();
  }

  async loadAuthAuditLogs(): Promise<void> {
    this.auditLoading = true;
    this.auditError = '';
    try {
      const qs = new URLSearchParams();
      qs.set('limit', '200');
      if (this.auditFilterEventType) qs.set('eventType', this.auditFilterEventType);
      if (this.auditFilterSuccess) qs.set('success', this.auditFilterSuccess);
      if (this.auditFilterFrom) qs.set('from', this.auditFilterFrom);
      if (this.auditFilterTo) qs.set('to', this.auditFilterTo);

      const rows = await this.apiJson<any[]>('GET', `/auth/audit?${qs.toString()}`);
      this.auditLogs = (rows || []).map((r) => ({
        id: Number(r.id),
        userId: r.userId == null ? null : Number(r.userId),
        username: String(r.username || ''),
        eventType: String(r.eventType || ''),
        success: Boolean(r.success),
        ip: String(r.ip || ''),
        userAgent: String(r.userAgent || ''),
        message: String(r.message || ''),
        createdAt: String(r.createdAt || ''),
        details: r.details && typeof r.details === 'object' ? r.details : {}
      }));
    } catch (e) {
      this.auditError = String((e as any)?.message || 'No se pudo cargar auditoría.');
    } finally {
      this.auditLoading = false;
    }
  }

  clearAuditFilters(): void {
    this.auditFilterEventType = '';
    this.auditFilterSuccess = '';
    this.auditFilterFrom = '';
    this.auditFilterTo = '';
    void this.loadAuthAuditLogs();
  }

  exportAuditCsv(): void {
    const rows = this.auditLogs || [];
    const esc = (v: unknown) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const headers = ['Fecha', 'Usuario', 'Evento', 'Exito', 'IP', 'Mensaje'];
    const lines = [headers.map(esc).join(',')];
    for (const r of rows) {
      lines.push(
        [
          r.createdAt,
          r.username,
          r.eventType,
          r.success ? 'SI' : 'NO',
          r.ip,
          r.message
        ]
          .map(esc)
          .join(',')
      );
    }
    const content = '\ufeff' + lines.join('\r\n') + '\r\n';
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `auth-audit-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  formatAuditIp(ip: string): string {
    const v = String(ip || '').trim();
    if (!v) return '-';
    if (v === '::1' || v === '127.0.0.1') return 'LOCALHOST';
    return v;
  }

  async saveUser(form: NgForm): Promise<void> {
    this.userMessage = '';
    this.userError = '';

    if (!form.valid) {
      this.setTransientNotice('userMessage', 'userError', 'error', 'Completa todos los campos.', 5000);
      return;
    }

    try {
      const expiresAtIso = this.toIsoFromLocalDatetime(this.userForm.expiresAt);
      if (this.userForm.isTemporary) {
        if (!expiresAtIso) {
          this.setTransientNotice('userMessage', 'userError', 'error', 'Selecciona la fecha y hora de expiración del usuario temporal.', 5000);
          return;
        }
        if (new Date(expiresAtIso).getTime() <= Date.now()) {
          this.setTransientNotice('userMessage', 'userError', 'error', 'La fecha/hora de expiración debe ser futura.', 5000);
          return;
        }
      }

      const payload = {
        username: this.userForm.username,
        email: this.userForm.email,
        role: this.userForm.role,
        status: this.userForm.status,
        isTemporary: Boolean(this.userForm.isTemporary),
        expiresAt: this.userForm.isTemporary ? expiresAtIso : null,
        ...(String(this.userForm.password || '').trim() ? { password: String(this.userForm.password) } : {})
      };

      if (this.editingUserId === null && !String(this.userForm.password || '').trim()) {
        this.setTransientNotice('userMessage', 'userError', 'error', 'La contraseña es obligatoria para crear usuario.', 5000);
        return;
      }

      const passwordIn = String(this.userForm.password || '').trim();
      if (passwordIn) {
        const policyErr = this.validateUserPasswordPolicy(passwordIn);
        if (policyErr) {
          this.setTransientNotice('userMessage', 'userError', 'error', policyErr, 5000);
          return;
        }
      }

      if (this.editingUserId === null) {
        await this.apiJson<any>('POST', '/users', payload);
        this.setTransientNotice('userMessage', 'userError', 'success', 'Usuario registrado correctamente.', 5000);
      } else {
        await this.apiJson<any>('PUT', `/users/${this.editingUserId}`, payload);
        this.setTransientNotice('userMessage', 'userError', 'success', 'Usuario actualizado correctamente.', 5000);
      }

      await this.loadSecurityFromApi();
      this.userForm = this.emptyUser();
      this.editingUserId = null;
      form.resetForm(this.userForm);
    } catch (e) {
      const code = String((e as any)?.apiCode || '');
      if (code === 'TEMP_READ_ONLY') {
        this.setTransientNotice('userMessage', 'userError', 'error', 'Tu cuenta temporal está en modo solo lectura y no puede guardar cambios.', 5000);
        return;
      }
      if (code === 'USER_EXPIRED') {
        this.setTransientNotice('userMessage', 'userError', 'error', 'Tu usuario temporal expiró. Solicita renovación al administrador.', 5000);
        return;
      }
      this.setTransientNotice('userMessage', 'userError', 'error', String((e as any)?.message || 'No se pudo guardar el usuario (backend).'), 5000);
    }
  }

  editUser(user: User): void {
    this.userForm = {
      ...user,
      isTemporary: Boolean(user.isTemporary),
      expiresAt: user.isTemporary ? this.toLocalDatetimeInput(user.expiresAt) : '',
      temporaryHours: null,
      password: ''
    };
    this.editingUserId = user.id;
    this.userMessage = '';
    this.userError = '';
    this.view = 'users-new';
  }

  async deleteUser(id: number): Promise<void> {
    try {
      await this.apiJson<any>('DELETE', `/users/${id}`);
      await this.loadSecurityFromApi();
      if (this.editingUserId === id) this.openNewUser();
    } catch {
      // Ignore.
    }
  }

  openUserResetPasswordModal(user: User): void {
    this.userResetPasswordModalOpen = true;
    this.userResetPasswordUserId = user.id;
    this.userResetPasswordUsername = user.username;
    this.userResetPasswordValue = '';
    this.userResetPasswordConfirm = '';
    this.userResetPasswordError = '';
  }

  closeUserResetPasswordModal(): void {
    this.userResetPasswordModalOpen = false;
    this.userResetPasswordUserId = null;
    this.userResetPasswordUsername = '';
    this.userResetPasswordValue = '';
    this.userResetPasswordConfirm = '';
    this.userResetPasswordError = '';
  }

  async confirmUserResetPassword(): Promise<void> {
    const id = this.userResetPasswordUserId;
    if (id == null) return;

    const pwd = String(this.userResetPasswordValue || '');
    const pwd2 = String(this.userResetPasswordConfirm || '');
    if (!pwd) {
      this.userResetPasswordError = 'Ingresa nueva contraseña.';
      return;
    }
    if (pwd !== pwd2) {
      this.userResetPasswordError = 'Las contraseñas no coinciden.';
      return;
    }
    const policyErr = this.validateUserPasswordPolicy(pwd);
    if (policyErr) {
      this.userResetPasswordError = policyErr;
      return;
    }

    try {
      await this.apiJson<any>('POST', `/users/${id}/reset-password`, { password: pwd });
      this.closeUserResetPasswordModal();
      this.pushToast('success', 'Contraseña restablecida correctamente.', 5000);
    } catch (e) {
      this.userResetPasswordError = String((e as any)?.message || 'No se pudo restablecer contraseña.');
    }
  }

  async saveRole(form: NgForm): Promise<void> {
    this.roleMessage = '';
    this.roleError = '';

    if (!form.valid) {
      this.setTransientNotice('roleMessage', 'roleError', 'error', 'Completa todos los campos.', 5000);
      return;
    }

    try {
      if (this.editingRoleId === null) {
        await this.apiJson<any>('POST', '/roles', {
          name: this.roleForm.name,
          description: this.roleForm.description
        });
        this.setTransientNotice('roleMessage', 'roleError', 'success', 'Rol registrado correctamente.', 5000);
      } else {
        await this.apiJson<any>('PUT', `/roles/${this.editingRoleId}`, {
          description: this.roleForm.description
        });
        this.setTransientNotice('roleMessage', 'roleError', 'success', 'Rol actualizado correctamente.', 5000);
      }

      await this.loadSecurityFromApi();
      this.roleForm = this.emptyRole();
      this.editingRoleId = null;
      form.resetForm(this.roleForm);
    } catch {
      this.setTransientNotice('roleMessage', 'roleError', 'error', 'No se pudo guardar el rol (backend).', 5000);
    }
  }

  editRole(role: Role): void {
    this.roleForm = { ...role };
    this.editingRoleId = role.id;
    this.roleMessage = '';
    this.roleError = '';
    this.view = 'roles-list';
  }

  async deleteRole(id: number): Promise<void> {
    try {
      await this.apiJson<any>('DELETE', `/roles/${id}`);
      await this.loadSecurityFromApi();
      if (this.editingRoleId === id) this.openNewRole();
    } catch {
      // Ignore.
    }
  }

  savePermissions(): void {
    this.roleMessage = 'Permisos guardados correctamente.';
    this.roleError = '';
  }

  private emptyUser(): User {
    return {
      id: 0,
      username: '',
      email: '',
      role: 'Operador',
      status: 'Activo',
      isTemporary: false,
      expiresAt: '',
      temporaryHours: null,
      createdAt: '',
      lastLogin: '',
      password: ''
    };
  }

  private emptyRole(): Role {
    return {
      id: 0,
      name: 'Operador',
      description: '',
      permissions: []
    };
  }

  get billingViewType(): BillingDocType {
    // Prefer the current form selection when available.
    if (this.billingForm?.type) return this.billingForm.type;
    if (this.view === 'billing-facturas') return 'Factura';
    if (this.view === 'billing-credit-notes') return 'Nota de crédito';
    return 'Boleta';
  }

  newClient(): void {
    this.clientForm = this.emptyClient();
    this.editingId = null;
    this.validationMessage = '';
    this.validationError = '';
    this.clientNameEditable = false;
    this.clientNameAutoFilled = false;
    this.clientNameManualOverride = false;
    this.clientAddressEditable = this.clientForm.documentType !== 'RUC';
    this.clientAddressAutoFilled = false;
    this.clientAddressManualOverride = false;
    this.clientPrevDocumentNumber = '';
    this.view = 'clients';
  }

  editClient(client: Client): void {
    this.clientForm = { ...client };
    this.editingId = client.id;
    this.validationMessage = '';
    this.validationError = '';
    // Editing existing data: allow manual adjustments.
    this.clientNameEditable = true;
    this.clientNameAutoFilled = false;
    this.clientNameManualOverride = false;
    this.clientAddressEditable = true;
    this.clientAddressAutoFilled = false;
    this.clientAddressManualOverride = false;
    this.clientPrevDocumentNumber = String(client.documentNumber || '').trim();
    this.view = 'clients';
  }

  cancelForm(): void {
    this.newClient();
  }

  async deleteClient(id: number): Promise<void> {
    try {
      await this.apiJson<{ ok: boolean }>('DELETE', `/clients/${id}`);
      await this.loadClientsFromApi();
      if (this.editingId === id) this.newClient();
    } catch {
      // Ignore.
    }
  }

  editVehicle(vehicle: Vehicle): void {
    this.vehicleForm = { ...vehicle };
    this.editingVehicleId = vehicle.id;
    this.vehicleMessage = '';
    this.vehicleError = '';
    this.view = 'vehicles-new';
  }

  async deleteVehicle(id: number): Promise<void> {
    try {
      await this.apiJson<{ ok: boolean }>('DELETE', `/vehicles/${id}`);
      await this.loadVehiclesFromApi();
      if (this.editingVehicleId === id) this.openNewVehicle();
    } catch {
      // Ignore.
    }
  }

  editDriver(driver: Driver): void {
    this.driverForm = { ...driver };
    this.editingDriverId = driver.id;
    this.driverMessage = '';
    this.driverError = '';
    this.resetDriverDniLookup();
    // Editing existing driver: allow manual edits.
    this.driverNameEditable = true;
    this.driverNameAutoFilled = false;
    this.driverNameManualOverride = false;
    this.driverPrevDni = String(driver.dni || '').trim();
    this.view = 'drivers-new';
  }

  async generateBillingDocument(form: NgForm): Promise<void> {
    this.billingMessage = '';
    this.billingError = '';
    this.billingLookupMessage = '';
    this.billingLookupError = '';

    if (!form.valid) {
      this.setTransientNotice('billingMessage', 'billingError', 'error', 'Completa los datos del comprobante.', 5000);
      return;
    }

    try {
      // Boleta rule: DNI required when Total >= 700.
      if (this.billingForm.type === 'Boleta') {
        const total = Number(this.billingForm.total);
        const doc = String(this.billingForm.clientDocument || '').trim();
        const name = String(this.billingForm.clientName || '').trim();

        if (Number.isFinite(total) && total >= 700) {
          if (!/^\d{8}$/.test(doc)) {
            this.billingError = 'Boleta: si el Total es S/ 700 o más, es obligatorio identificar al cliente con DNI (8 dígitos).';
            this.pushToast('error', this.billingError, 5000);
            return;
          }
          if (!name) {
            // Allow the user to type the name when lookup did not auto-fill.
            this.billingClientNameEditable = true;
            this.billingError = 'Boleta: si el Total es S/ 700 o más, es obligatorio ingresar el nombre del cliente.';
            this.pushToast('error', this.billingError, 5000);
            return;
          }
        } else {
          // For Total < 700, allow empty and default it.
          if (!doc) this.billingForm.clientDocument = '0';
          if (!name) this.billingForm.clientName = 'CLIENTE VARIOS';
        }
      }

      const created = await this.apiJson<{ id: number }>('POST', '/billing-documents', {
        type: this.billingForm.type,
        clientId: this.billingSelectedClientId,
        clientDocument:
          this.billingForm.type === 'Nota de crédito' && this.isCreditNoteReason('02')
            ? this.billingCreditNoteNewClientDocument
            : this.billingForm.clientDocument,
        clientName:
          this.billingForm.type === 'Nota de crédito' && this.isCreditNoteReason('02')
            ? this.billingCreditNoteNewClientName
            : this.billingForm.clientName,
        detail:
          this.billingForm.type === 'Nota de crédito' && this.isCreditNoteReason('03')
            ? this.billingCreditNoteCorrectedDetail
            : this.billingForm.detail,
        subtotal: Number(this.billingForm.subtotal),
        channel: this.billingForm.channel,
        certificatePath: this.billingForm.certificatePath,
        validateSunat: this.billingValidateSunat,
        affectedDocType: this.billingForm.type === 'Nota de crédito' ? this.billingAffectedDocType : undefined,
        affectedDocId: this.billingForm.type === 'Nota de crédito' ? this.billingAffectedDocId : undefined,
        creditNoteReasonCode: this.billingForm.type === 'Nota de crédito' ? this.billingCreditNoteReasonCode : undefined,
        creditNoteReason: this.billingForm.type === 'Nota de crédito' ? this.billingCreditNoteReason : undefined
      });

      await this.loadBillingFromApi();

      // SUNAT envío/validación se hace en backend antes de persistir.

      this.setTransientNotice(
        'billingMessage',
        'billingError',
        'success',
        this.billingForm.type === 'Nota de venta' && !this.billingValidateSunat
          ? 'Nota de venta generada. No se envi a SUNAT.'
          : 'Comprobante generado y guardado en la base de datos.',
        5000
      );
      this.billingForm = this.emptyBillingDocument(this.billingForm.type);
      this.editingBillingId = null;
      form.resetForm(this.billingForm);
    } catch (e) {
      const raw = String((e as any)?.message || '').trim();

      // Keep full detail in the UI (so user can copy it).
      this.billingMessage = '';
      this.billingError = raw || 'No se pudo generar el comprobante (backend).';

      const normalize = (s: string) => s.replace(/\s+/g, ' ').trim();
      const extractValor = (s: string) => {
        const m = s.match(/valor:\s*\"([^\"]+)\"/i) || s.match(/valor:\s*'([^']+)'/i);
        return m ? String(m[1]).trim() : '';
      };

      let friendly = '';

      if (/SERIES/i.test(raw) || /serie del comprobante/i.test(raw)) {
        friendly = 'Configura la serie del comprobante (Configuración -> Series) antes de emitir.';
      } else if (/CREDIT_NOTE_BAD_SERIES/i.test(raw) || /Serie de Nota de crédito inválida/i.test(raw)) {
        friendly = 'Nota de crédito: la serie configurada es inválida. Usa FC01 (si afecta Factura) o BC01 (si afecta Boleta) en Configuración -> Series.';
      } else if (/SUNAT rechaz[oó] el comprobante:/i.test(raw) || /SUNAT rechaz[oó]/i.test(raw)) {
        const sunatMsg = raw.replace(/^SUNAT rechaz[oó] el comprobante:\s*/i, '').trim();
        const v = extractValor(sunatMsg);

        if (/\b4093\b/.test(sunatMsg) || /ubigeo/i.test(sunatMsg)) {
          friendly = `SUNAT rechazó: Ubigeo inválido${v ? ` (${v})` : ''}. Corrige en Configuración -> Empresa -> Ubigeo.`;
        } else {
          friendly = `SUNAT rechazó: ${normalize(sunatMsg).slice(0, 180)}`;
        }
      } else if (/Valida configuraci[oó]n antes de emitir:/i.test(raw)) {
        friendly = normalize(raw.replace(/\n+/g, ' | ')).slice(0, 220);
      } else if (/Validation ZIP Filename error/i.test(raw) || /nombre del archivo ZIP es incorrecto/i.test(raw)) {
        friendly = 'SUNAT rechaza el nombre del ZIP. Para Nota de crédito, configura serie FC01 (Factura) o BC01 (Boleta) en Configuración -> Series.';
      } else if (/ubigeo/i.test(raw)) {
        const v = extractValor(raw);
        friendly = `Ubigeo inválido${v ? ` (${v})` : ''}. Corrige en Configuración -> Empresa -> Ubigeo.`;
      } else {
        friendly = raw ? normalize(raw).slice(0, 220) : 'No se pudo generar el comprobante (backend).';
      }

      // eslint-disable-next-line no-console
      console.error('Billing generate failed:', raw);
      this.pushToast('error', friendly, 5000);
    }
  }

  private async allocateBillingCode(type: BillingDocType): Promise<string | null> {
    // 1) Backend (BD): asignación atómica.
    try {
      const response = await fetch(`${this.apiBaseUrl}/document-series/allocate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ docType: type })
      });

      if (response.ok) {
        const data = (await response.json()) as { series: string; correlative: number };
        const corr = String(Number(data.correlative) || 0).padStart(4, '0');
        return `${data.series}-${corr}`;
      }
    } catch {
      // Sin backend o sin conexión.
    }

    // 2) Fallback local (demo): incrementa nextCorrelative en localStorage.
    try {
      const raw = localStorage.getItem('documentSeries');
      const list = raw ? (JSON.parse(raw) as DocumentSeriesConfig[]) : [];
      const cfg = list.find((x) => x.docType === type);
      if (!cfg) return null;

      const corr = Number(cfg.nextCorrelative) || 1;
      const updated = list.map((x) => (x.docType === type ? { ...x, nextCorrelative: corr + 1 } : x));
      localStorage.setItem('documentSeries', JSON.stringify(updated));

      return `${cfg.series}-${String(corr).padStart(4, '0')}`;
    } catch {
      return null;
    }
  }

  onBillingClientDocumentChange(): void {
    this.billingLookupMessage = '';
    this.billingLookupError = '';

    const doc = String(this.billingForm.clientDocument || '').trim();

    // If the user edits the document away from the selected client, detach the selection.
    if (this.billingSelectedClientId !== null && doc !== this.billingSelectedClientDocument) {
      this.billingSelectedClientId = null;
      this.billingSelectedClientDocument = '';
    }

    // Any document edit: clear stale auto-filled name immediately.
    if (doc !== this.billingPrevClientDocument) {
      if (this.billingClientNameAutoFilled && !this.billingClientNameManualOverride) {
        this.billingForm.clientName = '';
        this.billingClientNameAutoFilled = false;
      }

      // New doc: reset manual override.
      this.billingClientNameManualOverride = false;
      this.billingPrevClientDocument = doc;
    }
    // Only lookup when complete.
    if (doc.length !== 8 && doc.length !== 11) {
      this.cancelPendingBillingLookup();
      this.billingLastLookupKey = '';

      // Allow manual entry while doc is incomplete.
      this.billingClientNameEditable = true;
      this.billingClientNameAutoFilled = false;
      return;
    }

    const docType: DocumentType = doc.length === 11 ? 'RUC' : 'DNI';
    const localValidation = this.validateDocument(docType, doc);
    if (!localValidation.valid) {
      this.cancelPendingBillingLookup();
      return;
    }

    const key = `${docType}:${doc}`;
    if (key === this.billingLastLookupKey) return;
    this.billingLastLookupKey = key;

    // New doc: lock until lookup returns.
    this.billingClientNameEditable = false;
    this.billingClientNameAutoFilled = false;
    this.billingLastAutoFilledName = '';

    this.cancelPendingBillingLookup();
    this.billingLookupDebounceTimer = window.setTimeout(() => {
      this.billingLookupDebounceTimer = null;
      void this.lookupBillingClient(docType, doc);
    }, 450);
  }

  onBillingClientSelected(): void {
    const id = this.billingSelectedClientId;
    const c = id == null ? null : this.clients.find((x) => x.id === id) || null;
    if (!c) {
      this.billingSelectedClientDocument = '';
      return;
    }

    // Fill from existing client (no external lookup needed).
    this.cancelPendingBillingLookup();
    this.billingLookupMessage = '';
    this.billingLookupError = '';

    this.billingSelectedClientDocument = String(c.documentNumber || '').trim();
    this.billingForm.clientDocument = this.billingSelectedClientDocument;
    this.billingForm.clientName = String(c.name || '').trim();

    this.billingPrevClientDocument = this.billingSelectedClientDocument;
    this.billingLastLookupKey = `DB:${this.billingSelectedClientDocument}`;

    this.billingLastAutoFilledName = this.billingForm.clientName;
    this.billingClientNameAutoFilled = true;
    this.billingClientNameEditable = false;
    this.billingClientNameManualOverride = false;

    // If user selected a client, preselect the latest order for that client (optional UX).
    // Skip for Nota de crédito (it comes from the affected document).
    if (this.billingForm.type !== 'Nota de crédito') {
      const selectedOrder = this.billingSelectedOrderId == null ? null : this.orders.find((o) => o.id === this.billingSelectedOrderId) || null;
      if (selectedOrder && Number(selectedOrder.clientId) !== Number(c.id)) {
        this.billingSelectedOrderId = null;
      }

      if (this.billingSelectedOrderId == null) {
        const latest = this.getBillingOrderOptions()[0] || null;
        if (latest) {
          this.billingSelectedOrderId = latest.id;
          this.onBillingOrderSelected();
        }
      }
    }
  }

  private cancelPendingBillingLookup(): void {
    if (this.billingLookupDebounceTimer !== null) {
      clearTimeout(this.billingLookupDebounceTimer);
      this.billingLookupDebounceTimer = null;
    }
    if (this.billingLookupAbort) {
      this.billingLookupAbort.abort();
      this.billingLookupAbort = null;
    }
  }

  private async lookupBillingClient(docType: DocumentType, doc: string): Promise<void> {
    const key = `${docType}:${doc}`;
    this.billingLookupLoading = true;
    this.billingLookupAbort = new AbortController();

    try {
      const path = docType === 'RUC' ? `lookup/ruc/${encodeURIComponent(doc)}` : `lookup/dni/${encodeURIComponent(doc)}`;
      const resp = await fetch(`${this.apiBaseUrl}/${path}`, { signal: this.billingLookupAbort.signal });
      const data = (await resp.json().catch(() => null)) as any;

      const stillSame = this.billingLastLookupKey === key && String(this.billingForm.clientDocument || '').trim() === doc;
      if (!stillSame) return;

      if (!resp.ok || !data?.ok) {
        this.setTransientNotice(
          'billingLookupMessage',
          'billingLookupError',
          'error',
          data?.message || 'No se pudo consultar DNI/RUC.',
          5000
        );

        // Allow manual entry if lookup fails.
        this.billingClientNameEditable = true;
        this.billingClientNameAutoFilled = false;
        return;
      }

      const normalized = data.normalized as { name?: string };
      const name = String(normalized?.name || '').trim();

      // If upstream says ok but no data, treat as not found.
      if (!name) {
        // Clear any stale auto-filled value.
        if (
          this.billingClientNameAutoFilled &&
          this.billingLastAutoFilledName &&
          String(this.billingForm.clientName || '').trim() === this.billingLastAutoFilledName
        ) {
          this.billingForm.clientName = '';
        }
        this.setTransientNotice(
          'billingLookupMessage',
          'billingLookupError',
          'error',
          'Documento no encontrado. Puedes ingresar el cliente manualmente.',
          5000
        );
        this.billingClientNameEditable = true;
        this.billingClientNameAutoFilled = false;
        return;
      }

      if (normalized?.name) {
        const current = String(this.billingForm.clientName || '').trim();
        if (!current || current === this.billingLastAutoFilledName) {
          this.billingForm.clientName = name;
          this.billingLastAutoFilledName = name;
          this.billingClientNameAutoFilled = true;
          this.billingClientNameEditable = false;
        } else {
          this.billingClientNameAutoFilled = false;
          this.billingClientNameEditable = true;
        }
      }

      this.setTransientNotice(
        'billingLookupMessage',
        'billingLookupError',
        'success',
        'Datos del cliente rellenados correctamente.',
        5000
      );
    } catch (e) {
      if (String((e as any)?.name || '') === 'AbortError') return;
      this.setTransientNotice(
        'billingLookupMessage',
        'billingLookupError',
        'error',
        'No se pudo conectar para consultar DNI/RUC.',
        5000
      );

      // Allow manual entry if backend is down.
      this.billingClientNameEditable = true;
      this.billingClientNameAutoFilled = false;
    } finally {
      this.billingLookupLoading = false;
      this.billingLookupAbort = null;
    }
  }

  enableBillingClientNameEdit(): void {
    this.billingClientNameEditable = true;
    this.billingClientNameManualOverride = true;
    this.pushToast('success', 'Edicion de cliente habilitada.', 5000);
  }

  resetBillingForm(): void {
    this.billingForm = this.emptyBillingDocument(this.billingViewType);
    this.billingMessage = '';
    this.billingError = '';
    this.billingLookupMessage = '';
    this.billingLookupError = '';
    this.billingValidateSunat = false;
    this.billingAffectedDocType = 'Factura';
    this.billingAffectedDocId = '';
    this.billingCreditNoteReasonCode = '01';
    this.billingCreditNoteReason = '';
    this.billingAffectedBillingId = null;
    this.billingCreditNoteNewClientDocument = '';
    this.billingCreditNoteNewClientName = '';
    this.billingCreditNoteCorrectedDetail = '';
    this.billingSelectedClientId = null;
    this.billingSelectedClientDocument = '';
    this.billingSelectedOrderId = null;
    this.billingClientNameEditable = true;
    this.billingClientNameAutoFilled = false;
    this.billingClientNameManualOverride = false;
    this.billingLastAutoFilledName = '';
    this.billingPrevClientDocument = '';
    this.cancelPendingBillingLookup();
  }

  resetSunatConfig(): void {
    this.sunatEnvironment = 'Prueba';
    this.sunatEndpoint = '';
    this.sunatToken = '';
    this.sunatConnectionMessage = '';
    this.sunatConnectionError = '';
  }

  async sendBillingToSunat(id: number, channel: SunatChannel): Promise<void> {
    this.billingMessage = '';
    this.billingError = '';

    const endpoint = this.sunatEndpoint.trim();
    if (!endpoint) {
      await this.updateBillingInApi(id, { channel, sunatStatus: 'Borrador', response: 'Pendiente de envío (sin endpoint)' });
      this.billingError = 'Configura el endpoint SUNAT/OSE antes de enviar.';
      return;
    }

    // SUNAT SOAP cannot be called from the browser (CORS + SOAP). Use backend.
    if (/sunat\.gob\.pe/i.test(endpoint)) {
      try {
        const result = await this.apiJson<{ ok: boolean; sunatStatus: SunatStatus; response: string }>(
          'POST',
          `/billing-documents/${id}/send-sunat`,
          { channel }
        );

        await this.loadBillingFromApi();

        // eslint-disable-next-line no-console
        console.log('SUNAT send (backend)', result);

        if (!result.ok) {
          this.billingError = String(result.response || 'No se pudo enviar a SUNAT. Revisa configuración.');
          return;
        }

        this.billingMessage = result.sunatStatus === 'Aceptado' ? 'Envío a SUNAT completado.' : 'SUNAT rechazó el comprobante.';
      } catch (e) {
        // eslint-disable-next-line no-console
        console.error('SUNAT send failed:', e);
        this.billingError = String((e as any)?.message || 'No se pudo enviar a SUNAT (backend).');
      }
      return;
    }

    const doc = this.billingDocuments.find((item) => item.id === id);
    if (!doc) {
      this.billingError = 'No se encontró el documento en la lista. Recarga y vuelve a intentar.';
      return;
    }

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          ...(this.sunatToken.trim() ? { Authorization: `Bearer ${this.sunatToken.trim()}` } : {})
        },
        body: JSON.stringify({
          environment: this.sunatEnvironment,
          channel,
          document: doc
        })
      });

      const rawText = await response.text();
      // eslint-disable-next-line no-console
      console.log('SUNAT/OSE response', { status: response.status, ok: response.ok, body: rawText });
      let data: any = null;
      try {
        data = rawText ? JSON.parse(rawText) : null;
      } catch {
        data = null;
      }

      const acceptedFromPayload =
        typeof data?.accepted === 'boolean' ? data.accepted : typeof data?.estado === 'string' ? data.estado === 'ACEPTADO' : undefined;
      const accepted = acceptedFromPayload ?? response.ok;

      await this.updateBillingInApi(id, {
        channel,
        sunatStatus: accepted ? 'Aceptado' : 'Rechazado',
        response:
          String(data?.message || data?.mensaje || '').trim() ||
          (accepted ? 'CDR recibido y aceptado' : 'SUNAT rechazó el comprobante') ||
          String(rawText || '').trim()
      });

      this.billingMessage = accepted ? 'Envío a SUNAT completado.' : 'SUNAT rechazó el comprobante.';
    } catch {
      // No marcar como rechazado por fallas de conexión; queda pendiente para reintento.
      await this.updateBillingInApi(id, { channel, sunatStatus: 'Borrador', response: 'Pendiente de envío (error de conexión)' });
      this.billingError = 'No se pudo conectar con el endpoint SUNAT/OSE.';
    }
  }

  async rejectBilling(id: number): Promise<void> {
    await this.updateBillingInApi(id, { sunatStatus: 'Rechazado', response: 'SUNAT rechazó el comprobante' });
  }

  private async updateBillingInApi(
    id: number,
    patch: Partial<Pick<BillingDocument, 'sunatStatus' | 'response' | 'channel'>>
  ): Promise<void> {
    try {
      await this.apiJson<{ ok: boolean }>('PUT', `/billing-documents/${id}`, patch);
      await this.loadBillingFromApi();
    } catch {
      // Fallback (demo)
      this.billingDocuments = this.billingDocuments.map((d) => (d.id === id ? { ...d, ...patch } : d));
    }
  }

  editBillingDocument(document: BillingDocument): void {
    this.billingForm = { ...document };
    this.editingBillingId = document.id;
    this.billingMessage = '';
    this.billingError = '';
  }

  async testSunatConnection(): Promise<void> {
    this.sunatConnectionMessage = '';
    this.sunatConnectionError = '';

    if (!this.sunatEndpoint.trim()) {
      this.setTransientNotice(
        'sunatConnectionMessage',
        'sunatConnectionError',
        'error',
        'Configura primero el endpoint de prueba SUNAT/OSE.',
        5000
      );
      return;
    }

    try {
      const response = await fetch(this.sunatEndpoint.trim(), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          ...(this.sunatToken.trim() ? { Authorization: `Bearer ${this.sunatToken.trim()}` } : {})
        },
        body: JSON.stringify({ ping: true, environment: this.sunatEnvironment })
      });

      if (response.ok) {
        this.setTransientNotice(
          'sunatConnectionMessage',
          'sunatConnectionError',
          'success',
          'Conexión SUNAT/OSE de prueba OK.',
          5000
        );
      } else {
        this.setTransientNotice(
          'sunatConnectionMessage',
          'sunatConnectionError',
          'error',
          'El endpoint respondió con error. Revisa credenciales y URL de prueba.',
          5000
        );
      }
    } catch {
      this.setTransientNotice(
        'sunatConnectionMessage',
        'sunatConnectionError',
        'error',
        'No se pudo conectar al endpoint SUNAT/OSE.',
        5000
      );
    }
  }

  private async loadSunatConfigFromApi(): Promise<void> {
    try {
      const data = await this.apiJson<any>('GET', '/config/sunat');
      if (!data) return;
      this.sunatEnvironment = (String(data.environment || 'Prueba').toLowerCase().startsWith('prod') ? 'Producción' : 'Prueba') as any;
      this.sunatEndpoint = String(data.endpoint || this.sunatEndpoint);
      this.sunatToken = String(data.token || '');

      // Prefer explicit SOL fields (new DB columns). Fallback to token format.
      const solUser = String(data.solUser || '').trim();
      const solPass = String(data.solPass || '').trim();
      if (solUser || solPass) {
        this.sunatSolUser = solUser;
        this.sunatSolPass = solPass;
      } else {
        const token = this.sunatToken;
        if (token.includes('|')) {
          const i = token.indexOf('|');
          this.sunatSolUser = token.slice(0, i).trim();
          this.sunatSolPass = token.slice(i + 1).trim();
        }
      }
      this.sunatConfigLoaded = true;
    } catch {
      // Keep current values.
    }
  }

  onSunatSolCredsChange(): void {
    this.sunatSolMessage = '';
    this.sunatSolError = '';

    if (!this.sunatConfigLoaded) return;

    if (this.sunatSolDebounceTimer !== null) {
      clearTimeout(this.sunatSolDebounceTimer);
      this.sunatSolDebounceTimer = null;
    }

    this.sunatSolDebounceTimer = window.setTimeout(() => {
      this.sunatSolDebounceTimer = null;
      void this.saveSunatSolCreds();
    }, 600);
  }

  private async saveSunatSolCreds(): Promise<void> {
    const user = String(this.sunatSolUser || '').trim();
    const pass = String(this.sunatSolPass || '').trim();
    const token = user || pass ? `${user}|${pass}` : '';

    // Don't persist half-complete creds.
    if ((user && !pass) || (!user && pass)) {
      this.sunatSolError = 'Completa usuario y clave SOL.';
      return;
    }

    try {
      await this.apiJson<any>('PUT', '/config/sunat', {
        environment: this.sunatEnvironment,
        endpoint: this.sunatEndpoint,
        token,
        solUser: user,
        solPass: pass
      });
      this.sunatToken = token;
      this.sunatSolMessage = token ? 'Credenciales SOL guardadas.' : 'Credenciales SOL limpiadas.';
      this.pushToast('success', this.sunatSolMessage, 5000);
    } catch (e) {
      this.sunatSolError = String((e as any)?.message || 'No se pudo guardar credenciales SOL (backend).');
      this.pushToast('error', this.sunatSolError, 5000);
    }
  }

  onSunatConfigFieldChange(): void {
    this.sunatConfigMessage = '';
    this.sunatConfigError = '';
    if (!this.sunatConfigLoaded) return;

    if (this.sunatConfigDebounceTimer !== null) {
      clearTimeout(this.sunatConfigDebounceTimer);
      this.sunatConfigDebounceTimer = null;
    }

    this.sunatConfigDebounceTimer = window.setTimeout(() => {
      this.sunatConfigDebounceTimer = null;
      void this.saveSunatConfig();
    }, 600);
  }

  async saveSunatConfig(): Promise<void> {
    const solUser = String(this.sunatSolUser || '').trim();
    const solPass = String(this.sunatSolPass || '').trim();

    // Keep consistent validation with backend.
    if ((solUser && !solPass) || (!solUser && solPass)) {
      this.sunatConfigError = 'Completa usuario y clave SOL.';
      return;
    }

    try {
      await this.apiJson<any>('PUT', '/config/sunat', {
        environment: this.sunatEnvironment,
        endpoint: String(this.sunatEndpoint || '').trim(),
        token: String(this.sunatToken || ''),
        solUser,
        solPass
      });
      this.sunatConfigMessage = 'Configuración SUNAT guardada.';
    } catch (e) {
      this.sunatConfigError = String((e as any)?.message || 'No se pudo guardar configuración SUNAT (backend).');
    }
  }

  async deleteDriver(id: number): Promise<void> {
    try {
      await this.apiJson<{ ok: boolean }>('DELETE', `/drivers/${id}`);
      await this.loadDriversFromApi();
      if (this.editingDriverId === id) this.openNewDriver();
    } catch {
      // Ignore.
    }
  }

  cancelDriverForm(): void {
    this.openNewDriver();
  }

  private resetDriverDniLookup(): void {
    this.driverLookupLoading = false;
    this.driverLookupMessage = '';
    this.driverLookupError = '';
    this.driverDniFailCount = 0;
    this.driverLastLookupKey = '';
    this.driverNameAutoFilled = false;
    this.driverNameManualOverride = false;
    if (this.driverLookupDebounceTimer !== null) {
      clearTimeout(this.driverLookupDebounceTimer);
      this.driverLookupDebounceTimer = null;
    }
    if (this.driverLookupAbort) {
      this.driverLookupAbort.abort();
      this.driverLookupAbort = null;
    }
  }

  onDriverDniChange(): void {
    // Messages are shown as toasts.
    this.driverLookupMessage = '';
    this.driverLookupError = '';

    const dni = String(this.driverForm.dni || '').trim();

    // Keep it digits-only.
    const normalized = dni.replace(/\D+/g, '').slice(0, 8);
    if (normalized !== dni) this.driverForm.dni = normalized;

    // Any DNI edit: clear auto-filled name immediately.
    if (normalized !== this.driverPrevDni) {
      if (!this.driverNameManualOverride && this.driverNameAutoFilled) {
        this.driverForm.name = '';
        this.driverNameAutoFilled = false;
      }
      this.driverPrevDni = normalized;
    }

    // If DNI is incomplete/invalid, allow manual name entry.
    if (!/^\d{8}$/.test(normalized)) {
      this.cancelPendingDriverLookup();
      this.driverLastLookupKey = '';
      this.driverNameEditable = true;
      this.driverNameAutoFilled = false;
      this.driverNameManualOverride = false;
      return;
    }

    const key = `DNI:${normalized}`;
    if (key === this.driverLastLookupKey) return;

    // New DNI: auto-filled name was already cleared above.
    this.driverLastLookupKey = key;

    // New DNI: reset auto-fill/override state.
    this.driverNameAutoFilled = false;
    this.driverNameManualOverride = false;
    this.driverNameEditable = false;

    this.cancelPendingDriverLookup();
    this.driverLookupDebounceTimer = window.setTimeout(() => {
      this.driverLookupDebounceTimer = null;
      void this.lookupDriverByDni(normalized);
    }, 450);
  }

  private cancelPendingDriverLookup(): void {
    if (this.driverLookupDebounceTimer !== null) {
      clearTimeout(this.driverLookupDebounceTimer);
      this.driverLookupDebounceTimer = null;
    }
    if (this.driverLookupAbort) {
      this.driverLookupAbort.abort();
      this.driverLookupAbort = null;
    }
  }

  private async lookupDriverByDni(dni: string): Promise<void> {
    const key = `DNI:${dni}`;
    this.driverLookupLoading = true;
    this.driverLookupAbort = new AbortController();

    try {
      const resp = await fetch(`${this.apiBaseUrl}/lookup/dni/${encodeURIComponent(dni)}`, { signal: this.driverLookupAbort.signal });
      const data = (await resp.json().catch(() => null)) as any;

      const stillSame = this.driverLastLookupKey === key && String(this.driverForm.dni || '').trim() === dni;
      if (!stillSame) return;

      if (!resp.ok || !data?.ok) {
        this.driverDniFailCount += 1;
        // Immediately allow manual entry; no attempt counter in the UI.
        this.driverNameEditable = true;

        const base = data?.message || 'No se pudo consultar el DNI.';
        this.pushToast('error', `${base} Puedes escribir el nombre manualmente.`, 5000);
        return;
      }

      const name = String(data?.normalized?.name || '').trim();
      if (name) {
        if (!this.driverNameManualOverride) {
          this.driverForm.name = name;
          this.driverNameEditable = false;
        }
        this.driverNameAutoFilled = true;
        this.driverDniFailCount = 0;
        this.pushToast('success', 'Nombre rellenado por DNI.', 5000);
      } else {
        this.driverDniFailCount += 1;
        // Immediately allow manual entry; no attempt counter in the UI.
        this.driverNameEditable = true;
        this.pushToast('error', 'No se encontro nombre para el DNI. Puedes escribirlo manualmente.', 5000);
      }
    } catch (e) {
      if (String((e as any)?.name || '') === 'AbortError') return;
      this.driverDniFailCount += 1;
      this.driverNameEditable = true;
      this.pushToast('error', 'No se pudo conectar para consultar DNI. Puedes escribir el nombre manualmente.', 5000);
    } finally {
      this.driverLookupLoading = false;
      this.driverLookupAbort = null;
    }
  }

  enableDriverNameEdit(): void {
    this.driverNameEditable = true;
    this.driverNameManualOverride = true;
    this.pushToast('success', 'Edicion de nombre habilitada.', 5000);
  }

  async saveDriver(form: NgForm): Promise<void> {
    this.driverMessage = '';
    this.driverError = '';

    if (!form.valid) {
      this.setTransientNotice('driverMessage', 'driverError', 'error', 'Completa todos los campos del chofer.', 5000);
      return;
    }

    if (this.driverForm.dni.length !== 8) {
      this.setTransientNotice('driverMessage', 'driverError', 'error', 'El DNI debe tener 8 dígitos.', 5000);
      return;
    }

    try {
      const payload = {
        name: this.driverForm.name,
        dni: this.driverForm.dni,
        license: this.driverForm.license,
        phone: this.driverForm.phone,
        status: this.driverForm.status
      };

      if (this.editingDriverId === null) {
        await this.apiJson<Driver>('POST', '/drivers', payload);
        this.setTransientNotice('driverMessage', 'driverError', 'success', 'Chofer registrado correctamente.', 5000);
      } else {
        await this.apiJson<Driver>('PUT', `/drivers/${this.editingDriverId}`, payload);
        this.setTransientNotice('driverMessage', 'driverError', 'success', 'Chofer actualizado correctamente.', 5000);
      }

      await this.loadDriversFromApi();
      this.driverForm = this.emptyDriver();
      this.editingDriverId = null;
      form.resetForm(this.driverForm);
    } catch {
      this.setTransientNotice('driverMessage', 'driverError', 'error', 'No se pudo guardar el chofer (backend).', 5000);
    }
  }

  cancelVehicleForm(): void {
    this.openNewVehicle();
  }

  async saveVehicle(form: NgForm): Promise<void> {
    this.vehicleMessage = '';
    this.vehicleError = '';

    if (!form.valid) {
      this.setTransientNotice('vehicleMessage', 'vehicleError', 'error', 'Completa todos los campos del vehículo.', 5000);
      return;
    }

    try {
      const payload = {
        plate: this.vehicleForm.plate,
        capacity: Number(this.vehicleForm.capacity),
        status: this.vehicleForm.status,
        soat: this.vehicleForm.soat,
        technicalReview: this.vehicleForm.technicalReview
      };

      if (this.editingVehicleId === null) {
        await this.apiJson<Vehicle>('POST', '/vehicles', payload);
        this.setTransientNotice('vehicleMessage', 'vehicleError', 'success', 'Vehículo registrado correctamente.', 5000);
      } else {
        await this.apiJson<Vehicle>('PUT', `/vehicles/${this.editingVehicleId}`, payload);
        this.setTransientNotice('vehicleMessage', 'vehicleError', 'success', 'Vehículo actualizado correctamente.', 5000);
      }

      await this.loadVehiclesFromApi();
      this.vehicleForm = this.emptyVehicle();
      this.editingVehicleId = null;
      form.resetForm(this.vehicleForm);
    } catch {
      this.setTransientNotice('vehicleMessage', 'vehicleError', 'error', 'No se pudo guardar el vehículo (backend).', 5000);
    }
  }

  editOrder(order: Order): void {
    this.orderForm = { ...order, history: [...order.history] };
    this.editingOrderId = order.id;
    this.orderMessage = '';
    this.orderError = '';
    this.view = 'orders-new';
  }

  async deleteOrder(orderId: number): Promise<void> {
    try {
      await this.apiJson<{ ok: boolean }>('DELETE', `/orders/${orderId}`);
      await this.loadOrdersFromApi();
      await this.loadOrderHistoryFromApi();
      this.refreshRouteData();
    } catch {
      // Ignore.
    }
  }

  openAssignModal(orderId: number): void {
    const order = this.orders.find((item) => item.id === orderId);
    if (!order) return;

    this.assignModalOrderId = orderId;
    const vehicle = this.vehicles.find((v) => v.plate === order.vehicle);
    const driver = this.drivers.find((d) => d.name === order.driver);
    this.assignModalVehicleId = vehicle?.id ?? null;
    this.assignModalDriverId = driver?.id ?? null;
    this.assignModalError = '';
    this.assignModalOpen = true;
  }

  closeAssignModal(): void {
    this.assignModalOpen = false;
    this.assignModalOrderId = null;
    this.assignModalVehicleId = null;
    this.assignModalDriverId = null;
    this.assignModalError = '';
  }

  async confirmAssignModal(): Promise<void> {
    if (this.assignModalOrderId === null) return;
    if (this.assignModalVehicleId === null || this.assignModalDriverId === null) {
      this.assignModalError = 'Selecciona vehiculo y chofer.';
      return;
    }

    this.assignModalError = '';
    try {
      await this.apiJson<{ ok: boolean }>('POST', `/orders/${this.assignModalOrderId}/assign`, {
        vehicleId: this.assignModalVehicleId,
        driverId: this.assignModalDriverId
      });
      await this.loadOrdersFromApi();
      await this.loadOrderHistoryFromApi();
      this.refreshRouteData();
      this.closeAssignModal();
    } catch {
      this.assignModalError = 'No se pudo guardar la asignación (backend).';
    }
  }

  openAddressFixModal(order: Order): void {
    this.addressFixOrderId = order.id;
    this.addressFixMapsUrl = '';
    this.addressFixResolvedLabel = '';
    this.addressFixError = '';
    this.addressFixAddress = String(order.deliveryAddress || '').trim();
    this.addressFixOriginalAddress = this.addressFixAddress;
    this.addressFixResolving = false;
    this.addressFixSaving = false;
    this.addressFixModalOpen = true;
  }

  closeAddressFixModal(): void {
    this.addressFixModalOpen = false;
    this.addressFixOrderId = null;
    this.addressFixMapsUrl = '';
    this.addressFixAddress = '';
    this.addressFixOriginalAddress = '';
    this.addressFixResolvedLabel = '';
    this.addressFixError = '';
    this.addressFixResolving = false;
    this.addressFixSaving = false;
  }

  async resolveAddressFixFromMaps(): Promise<void> {
    const raw = String(this.addressFixMapsUrl || '').trim();
    this.addressFixError = '';
    this.addressFixResolvedLabel = '';

    if (!raw) {
      this.addressFixError = 'Pega un link de Google Maps (o escribe la direccion manualmente en "Direccion final").';
      return;
    }

    let token = this.getMapboxAccessToken();
    if (!token) {
      await this.loadPublicRuntimeConfig();
      token = this.getMapboxAccessToken();
    }
    if (!token) {
      this.addressFixError = 'Falta MAPBOX_PUBLIC_TOKEN en backend para poder obtener el nombre automaticamente. Puedes escribir la direccion manualmente.';
      return;
    }

    const parsed = await this.resolveGoogleMapsInput(raw);
    if (!parsed) {
      this.addressFixError = 'No se pudo leer el link. Pega un link de Google Maps valido o escribe la direccion manualmente.';
      return;
    }

    this.addressFixResolving = true;
    try {
      let label: string | null = null;

      if (parsed.kind === 'latlng') {
        label = await this.reverseGeocodeMapboxLatLng(parsed.lat, parsed.lng, token);
      } else if (parsed.kind === 'query') {
        const out = await this.geocodeMapboxAddress(parsed.query, token);
        label = out?.resolvedLabel ?? null;
      }

      if (!label) {
        this.addressFixError = 'No se encontro una direccion. Copia la direccion desde Google Maps y pegala en "Direccion final".';
        return;
      }

      this.addressFixResolvedLabel = label;

      const current = String(this.addressFixAddress || '').trim();
      if (!current || current === this.addressFixOriginalAddress) {
        this.addressFixAddress = label;
      }
    } catch {
      this.addressFixError = 'No se pudo consultar la direccion. Intenta con otra ubicacion o escribe la direccion manualmente.';
    } finally {
      this.addressFixResolving = false;
    }
  }

  async saveAddressFix(): Promise<void> {
    if (this.addressFixOrderId === null) return;

    const newAddress = String(this.addressFixAddress || '').trim();
    this.addressFixError = '';
    if (!newAddress) {
      this.addressFixError = 'La direccion final es requerida.';
      return;
    }

    this.addressFixSaving = true;
    try {
      await this.updateOrderAddress(this.addressFixOrderId, newAddress);
      this.closeAddressFixModal();
    } catch (e) {
      this.addressFixError = String((e as any)?.message || 'No se pudo guardar la direccion.');
    } finally {
      this.addressFixSaving = false;
    }
  }

  private async updateOrderAddress(orderId: number, deliveryAddress: string): Promise<void> {
    const order = this.orders.find((o) => o.id === orderId);
    if (!order) throw new Error('Pedido no encontrado.');

    const payload = {
      clientId: order.clientId,
      zone: order.zone,
      deliveryAddress,
      serviceType: order.serviceType,
      quantity: order.quantity,
      price: order.price,
      deliveryDate: order.deliveryDate,
      status: order.status
    };

    try {
      await this.apiJson<{ ok: boolean }>('PUT', `/orders/${orderId}`, payload);

      // Keep current view consistent.
      if (this.view === 'routes-daily') {
        // Force reload (same month) to reflect DB value immediately.
        this.routeOrdersLoadedMonth = '';
        await this.loadOrdersForRoutesView();
      } else {
        await this.loadOrdersFromApi();
      }

      await this.loadOrderHistoryFromApi();
    } catch (e) {
      // If the backend is offline, keep it working for the UI/map.
      const msg = String((e as any)?.message || e || '');
      const backendOffline = /failed to fetch|networkerror|fetch/i.test(msg);
      if (!backendOffline) throw e;
      this.orders = this.orders.map((o) => (o.id === orderId ? { ...o, deliveryAddress } : o));
    }

    // Keep route stats + map in sync.
    this.refreshRouteData();

    if (this.routeMapEnabled && this.routesMap) {
      await this.refreshRoutesMapMarkers();
    }

    this.pushToast('success', 'Direccion actualizada.', 3500);
  }

  getOrdersForSelectedDate(): Order[] {
    return this.orders.filter((order) => order.deliveryDate === this.routeDate);
  }

  getOrdersForSelectedMonth(): Order[] {
    const key = String(this.routeDate || '').slice(0, 7); // YYYY-MM
    if (key.length !== 7) return [];
    return this.orders.filter((order) => String(order.deliveryDate || '').startsWith(key));
  }

  getOrdersGroupedByZone(orders: Order[] = this.getOrdersForSelectedDate()): Array<{ zone: ZoneType; count: number }> {
    return this.zoneOptions.map((zone) => ({
      zone,
      count: orders.filter((order) => order.zone === zone).length
    }));
  }

  getRouteZoneMax(): number {
    return Math.max(1, this.orders.length);
  }

  refreshRouteData(): void {
    this.routeGroupByZone = this.getOrdersGroupedByZone();
    this.routeAssignments = this.orders
      .filter((order) => order.vehicle && order.driver)
      .map((order) => ({ orderId: order.id, vehicle: order.vehicle, driver: order.driver }));
  }

  toggleRouteMap(): void {
    this.routeMapEnabled = !this.routeMapEnabled;

    if (this.routeMapEnabled) {
      // Defer until the *ngIf map container is in the DOM.
      setTimeout(() => void this.initRoutesMap(), 0);
    } else {
      this.routeMapError = '';
      this.routesMapMarkers.forEach((m) => m.remove?.());
      this.routesMapMarkers = [];
      this.routesMap?.remove?.();
      this.routesMap = null;
    }
  }

  onRouteDateChange(): void {
    const date = String(this.routeDate || '').trim();
    if (!date) {
      this.refreshRouteData();
      if (this.routeMapEnabled) {
        this.refreshRoutesMapMarkers();
      }
      return;
    }

    // If month changed, reload data (so totals + table stay consistent).
    void this.loadOrdersForRoutesView().finally(() => {
      this.refreshRouteData();
      if (this.routeMapEnabled) {
        this.refreshRoutesMapMarkers();
      }
    });
  }

  private async initRoutesMap(): Promise<void> {
    this.routeMapError = '';

    try {
      await this.ensureMapboxLoaded();

      // If the map already exists (e.g. date change), just refresh markers.
      if (this.routesMap) {
        requestAnimationFrame(() => this.routesMap?.resize?.());
        await this.refreshRoutesMapMarkers();
        return;
      }

      // The container is rendered via *ngIf, so it may not exist yet.
      const canvas = await this.waitForElement('routes-map', 20);
      if (!canvas) {
        throw new Error('No se encontró el contenedor del mapa.');
      }

      // Default to Lima while we geocode per-order addresses.
      const center: [number, number] = [-77.0428, -12.0464];
      let token = this.getMapboxAccessToken();
      if (!token) {
        await this.loadPublicRuntimeConfig();
        token = this.getMapboxAccessToken();
      }
      if (!token) {
        throw new Error('Falta MAPBOX_PUBLIC_TOKEN en backend. Configúralo en variables de entorno.');
      }

      mapboxgl.accessToken = token;
      this.routesMap = new mapboxgl.Map({
        container: canvas,
        style: 'mapbox://styles/mapbox/streets-v12',
        center,
        zoom: 11,
        attributionControl: false
      });

      this.routesMap.addControl(new mapboxgl.NavigationControl({ visualizePitch: true }), 'bottom-right');

      await new Promise<void>((resolve) => this.routesMap.on('load', () => resolve()));
      requestAnimationFrame(() => this.routesMap?.resize?.());
      await this.refreshRoutesMapMarkers();
    } catch (err) {
      this.routeMapError = err instanceof Error ? err.message : 'No se pudo cargar el mapa.';
    }
  }

  private waitForElement(id: string, tries: number): Promise<HTMLElement | null> {
    return new Promise((resolve) => {
      let count = 0;
      const tick = () => {
        const el = document.getElementById(id);
        if (el) {
          resolve(el);
          return;
        }
        count += 1;
        if (count >= tries) {
          resolve(null);
          return;
        }
        setTimeout(tick, 50);
      };
      tick();
    });
  }

  private async refreshRoutesMapMarkers(): Promise<void> {
    if (!this.routesMap) return;

    // Clear existing markers.
    this.routesMapMarkers.forEach((m) => m.remove?.());
    this.routesMapMarkers = [];
    this.routesMapMarkerByOrderId.clear();

    const orders = this.getOrdersForSelectedDate();
    if (orders.length === 0) return;

    // Best-effort: geocode delivery addresses client-side.
    // If geocoding isn't available/allowed, fall back to zone centroids.
    const zoneCenters: Record<ZoneType, { lat: number; lng: number }> = {
      Centro: { lat: -12.0464, lng: -77.0428 },
      Norte: { lat: -12.0166, lng: -77.0580 },
      Sur: { lat: -12.1580, lng: -76.9903 },
      Este: { lat: -12.0350, lng: -76.9400 },
      Oeste: { lat: -12.0700, lng: -77.1200 }
    };

    const bounds = new mapboxgl.LngLatBounds();

    const token = this.getMapboxAccessToken();
    const canGeocode = !!token;

    const failed: number[] = [];
    const mismatched: number[] = [];

    for (const order of orders) {
      let pos: { lat: number; lng: number } | null = null;
      let resolvedLabel = '';
      let precise = false;

      if (canGeocode && order.deliveryAddress) {
        const key = order.deliveryAddress.trim();
        const cached = this.mapboxGeocodeCache.get(key);
        if (cached) {
          pos = { lng: cached.lng, lat: cached.lat };
          resolvedLabel = cached.resolvedLabel;
          precise = cached.precise;
        } else {
          try {
            const geocoded = await this.resolveDeliveryAddressLocation(key, token!);
            if (geocoded) {
              pos = { lng: geocoded.lng, lat: geocoded.lat };
              resolvedLabel = geocoded.resolvedLabel;
              precise = geocoded.precise;
              this.mapboxGeocodeCache.set(key, geocoded);
            }
          } catch {
            pos = null;
          }
        }
      }

      if (!pos) {
        pos = zoneCenters[order.zone];
        resolvedLabel = `Aprox. ${order.zone} (sin geocodificar)`;
        precise = false;
        failed.push(order.id);
      } else if (!precise) {
        // We found a referential location but it doesn't match the input address.
        mismatched.push(order.id);
      }

      const popupText = `Pedido #${order.id} · ${order.clientName}` +
        `\nDireccion (pedido): ${order.deliveryAddress || '-'}` +
        `\nUbicacion (mapa): ${resolvedLabel || '-'}`;

      const popup = new mapboxgl.Popup({ offset: 18 }).setText(popupText);
      const marker = new mapboxgl.Marker({ color: precise ? '#38bdf8' : '#a78bfa' })
        .setLngLat([pos.lng, pos.lat])
        .setPopup(popup)
        .addTo(this.routesMap);

      this.routesMapMarkers.push(marker);
      this.routesMapMarkerByOrderId.set(order.id, marker);
      bounds.extend([pos.lng, pos.lat]);
    }

    this.routesMap.fitBounds(bounds, { padding: 48, maxZoom: 14, duration: 600 });

    const parts: string[] = [];
    if (mismatched.length) {
      parts.push(`No coinciden con la dirección: ${mismatched.map((id) => `#${id}`).join(', ')} (se mostró referencia).`);
    }
    if (failed.length) {
      parts.push(`Sin geocodificar: ${failed.map((id) => `#${id}`).join(', ')} (se usó aproximación por zona).`);
    }
    this.routeMapError = parts.join(' ');
  }

  private parseGoogleMapsLink(
    input: string
  ):
    | { kind: 'latlng'; lat: number; lng: number }
    | { kind: 'query'; query: string }
    | null {
    const raw = String(input || '').trim();
    if (!raw) return null;

    // If the user pastes an address instead of a URL, treat it as a query.
    let url: URL | null = null;
    try {
      url = new URL(raw);
    } catch {
      return { kind: 'query', query: raw };
    }

    const href = url.href;

    // Common Google Maps patterns.
    // Prefer the pinned place coords (!3d/!4d) over the camera center (@lat,lng).
    // In many share links, the @ coords can be offset from the actual place.
    const bangRe = /!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/g;
    let lastBang: RegExpExecArray | null = null;
    for (;;) {
      const m = bangRe.exec(href);
      if (!m) break;
      lastBang = m;
    }
    if (lastBang) {
      const lat = Number(lastBang[1]);
      const lng = Number(lastBang[2]);
      if (Number.isFinite(lat) && Number.isFinite(lng)) return { kind: 'latlng', lat, lng };
    }

    const at = href.match(/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/);
    if (at) {
      const lat = Number(at[1]);
      const lng = Number(at[2]);
      if (Number.isFinite(lat) && Number.isFinite(lng)) return { kind: 'latlng', lat, lng };
    }

    const q = url.searchParams.get('q') || url.searchParams.get('query') || url.searchParams.get('destination');
    if (q) {
      const m = q.match(/^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$/);
      if (m) {
        const lat = Number(m[1]);
        const lng = Number(m[2]);
        if (Number.isFinite(lat) && Number.isFinite(lng)) return { kind: 'latlng', lat, lng };
      }
      return { kind: 'query', query: q };
    }

    const ll = url.searchParams.get('ll') || url.searchParams.get('center');
    if (ll) {
      const m = ll.match(/^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$/);
      if (m) {
        const lat = Number(m[1]);
        const lng = Number(m[2]);
        if (Number.isFinite(lat) && Number.isFinite(lng)) return { kind: 'latlng', lat, lng };
      }
    }

    // Last resort: try to use the full URL as a query.
    // Note: short links (maps.app.goo.gl) cannot be expanded client-side due to CORS.
    return null;
  }

  private async reverseGeocodeMapboxLatLng(lat: number, lng: number, token: string): Promise<string | null> {
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;

    const url =
      `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(String(lng))},${encodeURIComponent(String(lat))}.json` +
      `?access_token=${encodeURIComponent(token)}` +
      `&country=pe&language=es&limit=1&types=address`;

    const res = await fetch(url);
    if (!res.ok) return null;

    const data = (await res.json()) as { features?: Array<{ place_name?: string }> };
    const best = data.features?.[0];
    return best?.place_name ? String(best.place_name) : null;
  }

  private looksLikeUrl(value: string): boolean {
    return /^https?:\/\//i.test(String(value || '').trim());
  }

  private async resolveGoogleMapsInput(
    input: string
  ): Promise<{ kind: 'latlng'; lat: number; lng: number } | { kind: 'query'; query: string } | null> {
    const raw = String(input || '').trim();
    if (!raw) return null;

    const direct = this.parseGoogleMapsLink(raw);
    if (direct) return direct;

    if (!this.looksLikeUrl(raw)) return null;

    // Expand short links server-side.
    let host = '';
    try {
      host = new URL(raw).hostname.toLowerCase();
    } catch {
      return null;
    }

    const isGoogleShort = host === 'maps.app.goo.gl' || host === 'goo.gl' || host.endsWith('.goo.gl');
    if (!isGoogleShort) return null;

    try {
      const data = await this.apiJson<{
        ok: boolean;
        finalUrl?: string;
        parsed?: { kind: 'latlng'; lat: number; lng: number } | { kind: 'query'; query: string } | null;
      }>('POST', '/maps/resolve', { url: raw });

      if (data?.parsed) return data.parsed;
      if (data?.finalUrl) return this.parseGoogleMapsLink(String(data.finalUrl));
      return null;
    } catch {
      return null;
    }
  }

  private async resolveDeliveryAddressLocation(
    address: string,
    token: string
  ): Promise<{ lat: number; lng: number; resolvedLabel: string; precise: boolean } | null> {
    const raw = String(address || '').trim();
    if (!raw) return null;

    const parsed = await this.resolveGoogleMapsInput(raw);
    if (!parsed && this.looksLikeUrl(raw)) {
      // Don't try to geocode generic URLs.
      return null;
    }

    if (parsed?.kind === 'latlng') {
      const label = await this.reverseGeocodeMapboxLatLng(parsed.lat, parsed.lng, token);
      return {
        lat: parsed.lat,
        lng: parsed.lng,
        resolvedLabel: label || raw,
        precise: true
      };
    }

    if (parsed?.kind === 'query') {
      return await this.geocodeMapboxAddress(parsed.query, token);
    }

    // Default: treat as a normal address.
    return await this.geocodeMapboxAddress(raw, token);
  }

  private async geocodeMapboxAddress(
    address: string,
    token: string
  ): Promise<{ lat: number; lng: number; resolvedLabel: string; precise: boolean } | null> {
    // Bias to Lima/Perú for more consistent results.
    const limaProximity = '-77.0428,-12.0464';
    const limaBbox = '-77.20,-12.25,-76.80,-11.80';
    const query = `${address}, Lima, Peru`;

    const url =
      `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(query)}.json` +
      `?access_token=${encodeURIComponent(token)}` +
      `&country=pe&language=es&limit=10&autocomplete=false&types=address` +
      `&proximity=${encodeURIComponent(limaProximity)}` +
      `&bbox=${encodeURIComponent(limaBbox)}`;

    const res = await fetch(url);
    if (!res.ok) return null;

    const data = (await res.json()) as {
      features?: Array<{
        center?: [number, number];
        place_name?: string;
        relevance?: number;
        place_type?: string[];
        address?: string;
      }>;
    };

    const features = data.features ?? [];
    if (!features.length) return null;

    const wantedStreet = this.extractStreetKey(address);
    const wantedNumber = this.extractHouseNumber(address);

    const scored = features
      .filter((f) => Array.isArray(f.center) && f.center.length === 2)
      .map((f) => {
        const relevance = typeof f.relevance === 'number' ? f.relevance : 0;
        const isAddress = (f.place_type ?? []).includes('address');
        const hasHouseNumber = typeof f.address === 'string' && f.address.trim().length > 0;
        const normalizedPlace = this.normalizeForMatch(f.place_name ?? '');
        const streetMatch = wantedStreet ? normalizedPlace.includes(wantedStreet) : false;
        const numberMatch = wantedNumber
          ? new RegExp(`\\b${wantedNumber}\\b`).test(normalizedPlace)
          : false;

        // Prefer matches that include the same street token (and number when present).
        const score =
          relevance +
          (isAddress ? 0.2 : 0) +
          (hasHouseNumber ? 0.1 : 0) +
          (streetMatch ? 1.0 : 0) +
          (numberMatch ? 0.6 : 0);

        return { f, score, streetMatch, numberMatch };
      })
      .sort((a, b) => b.score - a.score);

    const bestStrict = scored.find((x) => x.streetMatch && (wantedNumber ? x.numberMatch : true))?.f;
    const bestLoose = scored.find((x) => x.streetMatch)?.f;
    const bestAny = scored[0]?.f;

    const best = bestStrict ?? bestLoose ?? bestAny;

    if (!best?.center) return null;

    // Only mark as precise if it matches the street token (and number if provided).
    const normalizedBest = this.normalizeForMatch(best.place_name ?? '');
    const precise = wantedStreet ? normalizedBest.includes(wantedStreet) : false;
    const preciseWithNumber = wantedNumber
      ? precise && new RegExp(`\\b${wantedNumber}\\b`).test(normalizedBest)
      : precise;

    return {
      lng: best.center[0],
      lat: best.center[1],
      resolvedLabel: best.place_name ?? address,
      precise: preciseWithNumber
    };
  }

  private normalizeForMatch(value: string): string {
    return value
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\b(av\.?|avenida)\b/g, 'av')
      .replace(/\b(jr\.?|jiron)\b/g, 'jr')
      .replace(/\b(ca\.?|calle)\b/g, 'calle')
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private extractHouseNumber(address: string): string {
    const m = address.match(/\b(\d{1,6})\b/);
    return m?.[1] ?? '';
  }

  private extractStreetKey(address: string): string {
    const normalized = this.normalizeForMatch(address);
    // Try to take street tokens before the number.
    const number = this.extractHouseNumber(normalized);
    const before = number ? normalized.split(new RegExp(`\\b${number}\\b`))[0] : normalized;
    const tokens = before.split(' ').filter(Boolean);
    // Remove generic tokens.
    const filtered = tokens.filter((t) => !['av', 'jr', 'calle', 'lima', 'peru'].includes(t));
    // Keep first 1-2 meaningful tokens as a key.
    return filtered.slice(0, 2).join(' ');
  }

  focusRouteOrder(orderId: number): void {
    if (!this.routesMap) return;
    const marker = this.routesMapMarkerByOrderId.get(orderId);
    if (!marker) return;

    const ll = marker.getLngLat?.();
    if (!ll) return;
    const currentZoom = this.routesMap.getZoom?.() ?? 12;
    this.routesMap.flyTo({ center: [ll.lng, ll.lat], zoom: Math.max(currentZoom, 13), duration: 600 });
    marker.getPopup?.()?.addTo(this.routesMap);
  }

  private async ensureMapboxLoaded(): Promise<void> {
    if (typeof window === 'undefined') return;
    if ((window as any).mapboxgl) return;

    if (this.mapboxLoading) return this.mapboxLoading;

    // We load Mapbox via CDN in index.html. If it's missing, explain clearly.
    this.mapboxLoading = new Promise<void>((resolve, reject) => {
      const tries = 40;
      let count = 0;
      const tick = () => {
        count += 1;
        if ((window as any).mapboxgl) {
          resolve();
          return;
        }
        if (count >= tries) {
          reject(new Error('Mapbox no cargó. Revisa los <script>/<link> en src/index.html.'));
          return;
        }
        setTimeout(tick, 50);
      };
      tick();
    });

    return this.mapboxLoading;
  }

  async assignRoute(): Promise<void> {
    if (this.routeAssignmentOrderId === null || !this.routeAssignmentVehicle || !this.routeAssignmentDriver) {
      return;
    }

    const vehicle = this.vehicles.find((v) => v.plate === this.routeAssignmentVehicle);
    const driver = this.drivers.find((d) => d.name === this.routeAssignmentDriver);
    if (!vehicle || !driver) return;

    try {
      await this.apiJson<{ ok: boolean }>('POST', `/orders/${this.routeAssignmentOrderId}/assign`, {
        vehicleId: vehicle.id,
        driverId: driver.id
      });
      await this.loadOrdersFromApi();
      await this.loadOrderHistoryFromApi();
      this.refreshRouteData();
    } catch {
      // Ignore.
    }
  }

  cancelOrderForm(): void {
    this.openNewOrder();
  }

  async saveOrder(form: NgForm): Promise<void> {
    this.orderMessage = '';
    this.orderError = '';

    if (!form.valid) {
      this.setTransientNotice('orderMessage', 'orderError', 'error', 'Completa todos los campos del pedido.', 5000);
      return;
    }

    const selectedClientId = Number(this.orderForm.clientId);
    const client = this.clients.find((item) => Number(item.id) === selectedClientId);
    if (!client) {
      this.setTransientNotice('orderMessage', 'orderError', 'error', 'Selecciona un cliente válido.', 5000);
      return;
    }

    const payload = {
      clientId: client.id,
      zone: this.orderForm.zone,
      deliveryAddress: this.orderForm.deliveryAddress,
      serviceType: this.orderForm.serviceType,
      quantity: Number(this.orderForm.quantity),
      price: Number(this.orderForm.price),
      deliveryDate: this.orderForm.deliveryDate,
      status: this.orderForm.status
    };

    try {
      let orderId: number;
      if (this.editingOrderId === null) {
        const created = await this.apiJson<{ id: number }>('POST', '/orders', payload);
        orderId = Number(created.id);
        this.setTransientNotice('orderMessage', 'orderError', 'success', 'Pedido generado correctamente.', 5000);
      } else {
        await this.apiJson<{ ok: boolean }>('PUT', `/orders/${this.editingOrderId}`, payload);
        orderId = this.editingOrderId;
        this.setTransientNotice('orderMessage', 'orderError', 'success', 'Pedido actualizado correctamente.', 5000);
      }

      // Si el usuario seleccionó vehículo/chofer, registra asignación.
      const plate = String(this.orderForm.vehicle || '').trim();
      const driverName = String(this.orderForm.driver || '').trim();
      if (plate && driverName) {
        const vehicle = this.vehicles.find((v) => v.plate === plate);
        const driver = this.drivers.find((d) => d.name === driverName);
        if (vehicle && driver) {
          await this.apiJson<{ ok: boolean }>('POST', `/orders/${orderId}/assign`, {
            vehicleId: vehicle.id,
            driverId: driver.id
          });
        }
      }

      await this.loadOrdersFromApi();
      await this.loadOrderHistoryFromApi();
      this.refreshRouteData();

      this.orderForm = this.emptyOrder();
      this.editingOrderId = null;
      form.resetForm(this.orderForm);
      this.view = 'orders-list';
    } catch {
      this.setTransientNotice('orderMessage', 'orderError', 'error', 'No se pudo guardar el pedido (backend).', 5000);
    }
  }

  assignVehicleAndDriver(orderId: number): void {
    // Legacy handler used by the action button.
    // ERP behavior: user chooses vehicle/driver (not random).
    this.openAssignModal(orderId);
  }

  async changeOrderStatus(orderId: number): Promise<void> {
    const order = this.orders.find((o) => o.id === orderId);
    if (!order) return;
    const nextStatus = order.status === 'Pendiente' ? 'En ruta' : order.status === 'En ruta' ? 'Entregado' : 'Pendiente';

    try {
      await this.apiJson<{ ok: boolean }>('POST', `/orders/${orderId}/status`, { status: nextStatus });
      await this.loadOrdersFromApi();
      await this.loadOrderHistoryFromApi();
      this.refreshRouteData();
    } catch {
      // Ignore.
    }
  }

  getClientName(clientId: number): string {
    return this.clients.find((client) => client.id === clientId)?.name ?? 'Cliente no encontrado';
  }

  async saveClient(form: NgForm): Promise<void> {
    this.validationMessage = '';
    this.validationError = '';
    this.lookupMessage = '';
    this.lookupError = '';

    if (!form.valid) {
      this.setClientTransientMessage('validationError', 'Completa los campos principales.', 5000);
      return;
    }

    const localValidation = this.validateDocument(this.clientForm.documentType, this.clientForm.documentNumber);
    if (!localValidation.valid) {
      this.setClientTransientMessage('validationError', localValidation.message, 5000);
      return;
    }

    try {
      const payload = {
        documentType: this.clientForm.documentType,
        documentNumber: this.clientForm.documentNumber,
        name: this.clientForm.name,
        address: this.clientForm.address || '',
        phone: this.clientForm.phone || '',
        email: this.clientForm.email || '',
        clientType: this.clientForm.clientType
      };

      if (this.editingId === null) {
        await this.apiJson<Client>('POST', '/clients', payload);
        this.setClientTransientMessage('validationMessage', 'Cliente registrado correctamente.', 5000);
      } else {
        await this.apiJson<Client>('PUT', `/clients/${this.editingId}`, payload);
        this.setClientTransientMessage('validationMessage', 'Cliente actualizado correctamente.', 5000);
      }

      await this.loadClientsFromApi();
    } catch {
      this.setClientTransientMessage('validationError', 'No se pudo guardar el cliente (backend).', 5000);
      return;
    }

    this.clientForm = this.emptyClient();
    this.editingId = null;
    form.resetForm(this.clientForm);
  }

  onClientTypeChange(): void {
    this.clientForm.documentType = this.clientForm.clientType === 'empresa' ? 'RUC' : 'DNI';
    this.validationError = '';
    this.validationMessage = '';
    this.lookupMessage = '';
    this.lookupError = '';
    this.cancelPendingLookup();

    // Reset auto-fill state.
    this.lastLookupKey = '';
    this.lastAutoFilledName = '';
    this.lastAutoFilledAddress = '';
    this.clientNameAutoFilled = false;
    this.clientAddressAutoFilled = false;
    this.clientNameEditable = false;
    this.clientAddressEditable = this.clientForm.documentType !== 'RUC';
    this.clientNameManualOverride = false;
    this.clientAddressManualOverride = false;
    this.clientPrevDocumentNumber = '';

    // If the user already typed a document number, re-run lookup
    // using the new expected document type/length.
    this.onClientDocumentChange();
  }

  private setClientTransientMessage(
    field: 'validationMessage' | 'validationError',
    value: string,
    ms = 5000
  ): void {
    // Reset both client status messages for a clean ERP feel.
    this.validationMessage = '';
    this.validationError = '';

    (this as any)[field] = value;

    this.pushToast(field === 'validationMessage' ? 'success' : 'error', value, ms);

    if (this.clientMessageTimer !== null) {
      clearTimeout(this.clientMessageTimer);
      this.clientMessageTimer = null;
    }

    this.clientMessageTimer = window.setTimeout(() => {
      // Only clear if message wasn't replaced.
      if ((this as any)[field] === value) (this as any)[field] = '';
      this.clientMessageTimer = null;
    }, ms);
  }

  private setLookupTransientMessage(field: 'lookupMessage' | 'lookupError', value: string, ms = 5000): void {
    this.lookupMessage = '';
    this.lookupError = '';

    (this as any)[field] = value;

    this.pushToast(field === 'lookupMessage' ? 'success' : 'error', value, ms);

    if (this.lookupMessageTimer !== null) {
      clearTimeout(this.lookupMessageTimer);
      this.lookupMessageTimer = null;
    }

    this.lookupMessageTimer = window.setTimeout(() => {
      if ((this as any)[field] === value) (this as any)[field] = '';
      this.lookupMessageTimer = null;
    }, ms);
  }

  onClientDocumentChange(): void {
    // Auto lookup only when the document is complete/valid.
    this.lookupMessage = '';
    this.lookupError = '';

    const docType = this.clientForm.documentType;
    const doc = this.clientForm.documentNumber.trim();

    // Any document edit: clear auto-filled fields immediately (avoid stale mismatch).
    if (doc !== this.clientPrevDocumentNumber) {
      if (!this.clientNameManualOverride && this.clientNameAutoFilled) {
        this.clientForm.name = '';
        this.clientNameAutoFilled = false;
      }
      if (!this.clientAddressManualOverride && this.clientAddressAutoFilled) {
        this.clientForm.address = '';
        this.clientAddressAutoFilled = false;
      }

      // New document: reset manual overrides.
      this.clientNameManualOverride = false;
      this.clientAddressManualOverride = false;

      this.clientPrevDocumentNumber = doc;
    }

    const expectedLen = docType === 'RUC' ? 11 : 8;
    if (doc.length < expectedLen) {
      this.cancelPendingLookup();
      this.lastLookupKey = '';

      // Allow manual entry while the document is incomplete.
      this.clientNameEditable = true;
      this.clientAddressEditable = true;
      this.clientNameAutoFilled = false;
      this.clientAddressAutoFilled = false;
      return;
    }

    // Only lookup when it fully matches our local validation.
    const localValidation = this.validateDocument(docType, doc);
    if (!localValidation.valid) {
      this.cancelPendingLookup();

      // Invalid doc: allow manual entry.
      this.clientNameEditable = true;
      this.clientAddressEditable = true;
      this.clientNameAutoFilled = false;
      this.clientAddressAutoFilled = false;
      return;
    }

    const key = `${docType}:${doc}`;
    if (key === this.lastLookupKey) return;

    // New document: lock fields until lookup returns (ERP behavior).
    this.clientNameEditable = false;
    this.clientNameAutoFilled = false;
    this.clientAddressEditable = docType !== 'RUC';
    this.clientAddressAutoFilled = false;
    // Keep lastAutoFilled* values for immediate clearing logic when the doc changes.

    this.cancelPendingLookup();
    this.lookupDebounceTimer = window.setTimeout(() => {
      this.lookupDebounceTimer = null;
      void this.lookupClientByDocumentInternal(docType, doc);
    }, 450);
  }

  private cancelPendingLookup(): void {
    if (this.lookupDebounceTimer !== null) {
      clearTimeout(this.lookupDebounceTimer);
      this.lookupDebounceTimer = null;
    }
    if (this.lookupAbort) {
      this.lookupAbort.abort();
      this.lookupAbort = null;
    }
  }

  private async lookupClientByDocumentInternal(docType: DocumentType, doc: string): Promise<void> {
    const key = `${docType}:${doc}`;
    this.lastLookupKey = key;

    const path = docType === 'RUC' ? `lookup/ruc/${encodeURIComponent(doc)}` : `lookup/dni/${encodeURIComponent(doc)}`;
    this.lookupLoading = true;
    this.lookupAbort = new AbortController();

    try {
      const resp = await fetch(`${this.apiBaseUrl}/${path}`, { signal: this.lookupAbort.signal });
      const data = (await resp.json().catch(() => null)) as any;

      if (!resp.ok || !data?.ok) {
        // Only show error if the user still has the same complete document.
        const stillSame = this.lastLookupKey === key && this.clientForm.documentNumber.trim() === doc;
        if (stillSame) {
          this.setLookupTransientMessage(
            'lookupError',
            data?.message || 'No se pudo consultar el documento. Revisa el servicio.',
            5000
          );

          // Allow manual entry if lookup fails.
          this.clientNameEditable = true;
          this.clientAddressEditable = true;
          this.clientNameAutoFilled = false;
          this.clientAddressAutoFilled = false;
        }
        return;
      }

      const normalized = data.normalized as { name?: string; address?: string };
      const stillSame = this.lastLookupKey === key && this.clientForm.documentNumber.trim() === doc;
      if (!stillSame) return;

      const normalizedName = String(normalized?.name || '').trim();
      const normalizedAddress = String(normalized?.address || '').trim();

      // If the upstream responds ok but without useful data, treat it as not found.
      if (!normalizedName) {
        // Clear any stale auto-filled values.
        if (this.clientNameAutoFilled && this.lastAutoFilledName && this.clientForm.name.trim() === this.lastAutoFilledName) {
          this.clientForm.name = '';
        }
        if (
          this.clientAddressAutoFilled &&
          this.lastAutoFilledAddress &&
          this.clientForm.address.trim() === this.lastAutoFilledAddress
        ) {
          this.clientForm.address = '';
        }
        this.setLookupTransientMessage('lookupError', 'Documento no encontrado. Puedes ingresar los datos manualmente.', 5000);
        this.clientNameEditable = true;
        this.clientAddressEditable = true;
        this.clientNameAutoFilled = false;
        this.clientAddressAutoFilled = false;
        return;
      }

      // Avoid overwriting manual edits: only overwrite if empty or previously auto-filled.
      if (normalizedName) {
        const current = this.clientForm.name.trim();
        if (!current || current === this.lastAutoFilledName) {
          this.clientForm.name = normalizedName;
          this.lastAutoFilledName = normalizedName;
          this.clientNameAutoFilled = true;
          this.clientNameEditable = false;
        } else {
          this.clientNameEditable = true;
          this.clientNameAutoFilled = false;
        }
      }
      if (docType === 'RUC' && normalizedAddress) {
        const current = this.clientForm.address.trim();
        if (!current || current === this.lastAutoFilledAddress) {
          this.clientForm.address = normalizedAddress;
          this.lastAutoFilledAddress = normalizedAddress;
          this.clientAddressAutoFilled = true;
          this.clientAddressEditable = false;
        } else {
          this.clientAddressEditable = true;
          this.clientAddressAutoFilled = false;
        }
      }

      // If DNI, keep address editable.
      if (docType === 'DNI') {
        this.clientAddressEditable = true;
        this.clientAddressAutoFilled = false;
      }

      this.setLookupTransientMessage('lookupMessage', 'Datos encontrados y rellenados correctamente.', 5000);
    } catch (e) {
      if (String((e as any)?.name || '') === 'AbortError') return;
      this.setLookupTransientMessage('lookupError', 'No se pudo conectar con el backend para consultar DNI/RUC.', 5000);

      // Allow manual entry if backend is down.
      this.clientNameEditable = true;
      this.clientAddressEditable = true;
      this.clientNameAutoFilled = false;
      this.clientAddressAutoFilled = false;
    } finally {
      this.lookupLoading = false;
      this.lookupAbort = null;
    }
  }

  enableClientNameEdit(): void {
    this.clientNameEditable = true;
    this.clientNameManualOverride = true;
    this.pushToast('success', 'Edicion de nombre habilitada.', 5000);
  }

  enableClientAddressEdit(): void {
    this.clientAddressEditable = true;
    this.clientAddressManualOverride = true;
    this.pushToast('success', 'Edicion de direccion habilitada.', 5000);
  }

  get documentLabel(): string {
    return this.clientForm.documentType === 'RUC' ? 'RUC' : 'DNI';
  }

  private emptyClient(): Client {
    return {
      id: 0,
      documentType: 'RUC',
      documentNumber: '',
      name: '',
      address: '',
      phone: '',
      email: '',
      clientType: 'empresa'
    };
  }

  private emptyOrder(): Order {
    return {
      id: 0,
      clientId: null,
      clientName: '',
      zone: 'Centro',
      deliveryAddress: '',
      serviceType: 'Agua potable',
      quantity: null,
      price: null,
      deliveryDate: '',
      status: 'Pendiente',
      vehicle: '',
      driver: '',
      history: []
    };
  }

  private emptyVehicle(): Vehicle {
    return {
      id: 0,
      plate: '',
      capacity: 0,
      status: 'Disponible',
      soat: '',
      technicalReview: ''
    };
  }

  private emptyDriver(): Driver {
    return {
      id: 0,
      name: '',
      dni: '',
      license: '',
      phone: '',
      status: 'Activo'
    };
  }

  private emptyBillingDocument(type: BillingDocType = 'Boleta'): BillingDocument {
    return {
      id: 0,
      type,
      clientDocument: '',
      clientName: '',
      detail: '',
      subtotal: 0,
      igv: 0,
      total: 0,
      // These paths are assigned by the backend when a document is generated.
      xmlPath: '',
      zipPath: '',
      pdfPath: '',
      cdrPath: '',
      certificatePath: '/certificados/firma.pfx',
      channel: 'API directa SUNAT',
      sunatStatus: 'Borrador',
      response: ''
    };
  }

  private openBillingView(view: 'billing-boletas' | 'billing-facturas' | 'billing-credit-notes', type: BillingDocType): void {
    this.billingForm = this.emptyBillingDocument(type);
    // Pull current configured certificate path (if any).
    const cert = String(this.certificateConfig?.path || '').trim();
    if (cert) this.billingForm.certificatePath = cert;
    this.billingValidateSunat = false;
    this.editingBillingId = null;
    this.billingMessage = '';
    this.billingError = '';
    this.billingLookupMessage = '';
    this.billingLookupError = '';
    this.cancelPendingBillingLookup();
    this.view = view;
    void this.loadClientsFromApi();
    void this.loadOrdersFromApi();
    void this.loadBillingFromApi();
  }

  private refreshOrderHistory(): void {
    this.orderHistory = this.orders.flatMap((order) =>
      order.history.map((message) => ({ id: order.id, message: `Pedido #${order.id}: ${message}` }))
    );
  }

  private async loadOrderHistoryFromApi(): Promise<void> {
    try {
      const entries = await Promise.all(
        this.orders.map(async (o) => {
          const events = await this.apiJson<Array<{ message: string }>>('GET', `/orders/${o.id}/events`);
          return { orderId: o.id, messages: events.map((e) => String(e.message ?? '')) };
        })
      );

      const byId = new Map<number, string[]>();
      entries.forEach((e) => byId.set(e.orderId, e.messages.filter(Boolean)));
      this.orders = this.orders.map((o) => ({ ...o, history: byId.get(o.id) ?? [] }));
      this.refreshOrderHistory();
    } catch {
      // Fallback: keep whatever is already in memory.
      this.refreshOrderHistory();
    }
  }

  private getTodayDate(): string {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  private validateDocument(documentType: DocumentType, value: string): { valid: boolean; message: string } {
    const normalized = value.trim();

    if (documentType === 'DNI') {
      return /^\d{8}$/.test(normalized)
        ? { valid: true, message: '' }
        : { valid: false, message: 'El DNI debe tener 8 dígitos.' };
    }

    if (!/^\d{11}$/.test(normalized)) {
      return { valid: false, message: 'El RUC debe tener 11 dígitos.' };
    }

    if (!this.isValidRuc(normalized)) {
      return { valid: false, message: 'El RUC no es válido.' };
    }

    return { valid: true, message: '' };
  }

  billingBoletaRequiresDni(): boolean {
    if (this.billingForm?.type !== 'Boleta') return false;
    const total = Number(this.billingForm.total);
    return Number.isFinite(total) && total >= 700;
  }

  billingClientDocumentIsRequired(): boolean {
    // Factura/Nota de venta/Nota de crédito: always required.
    // Boleta: required only when Total >= 700.
    return this.billingForm?.type === 'Boleta' ? this.billingBoletaRequiresDni() : true;
  }

  billingClientNameIsRequired(): boolean {
    // Name is required except Boleta with Total < 700 (allowed; backend will default to CLIENTE VARIOS).
    return this.billingForm?.type === 'Boleta' ? this.billingBoletaRequiresDni() : true;
  }

  private isValidRuc(ruc: string): boolean {
    const weights = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
    const digits = ruc.split('').map((digit) => Number(digit));
    const sum = weights.reduce((total, weight, index) => total + weight * digits[index], 0);
    const remainder = 11 - (sum % 11);
    const checkDigit = remainder === 10 ? 0 : remainder === 11 ? 1 : remainder;
    return checkDigit === digits[10];
  }

  // External validation is handled by our backend lookup endpoints.
}
