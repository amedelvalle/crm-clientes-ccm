import { useEffect, useMemo, useState } from 'react'
import { supabase } from './lib/supabase'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import logoCcm from './assets/logo-ccm.png'

const CLIENT_BATCH_SIZE = 500
const TX_BATCH_SIZE = 250
const PAGE_SIZE = 25
const EXCLUDED_EXPEDIENTES = new Set(['3864'])

const DEFAULT_FILTERS = {
  search: '',
  tipoCliente: 'PRIVADO',
  estado: 'todos',
  clasificacion: 'todos',
  anioComparativo: '',
  tipoClienteComparativo: 'PRIVADO'
}

const TABS = [
  { key: 'dashboard', label: 'Dashboard' },
  { key: 'clientes', label: 'Clientes' },
  { key: 'comparativo', label: 'Comparativo anual' },
  { key: 'exportar', label: 'Exportar' },
  { key: 'carga', label: 'Carga' },
  { key: 'historial', label: 'Historial' },
  { key: 'logica', label: 'Lógica del reporte' }
]

const RELATION_ORDER = {
  nuevo_reciente: 1,
  nuevo_en_seguimiento: 2,
  nuevo_sin_recompra: 3,
  activo: 4,
  vigilancia: 5,
  en_riesgo: 6,
  inactivo_reciente: 7,
  inactivo_prolongado: 8,
  recuperado: 9,
  inactivo_forzado: 10
}

const STRATEGY_ORDER = {
  bronce: 1,
  plata: 2,
  oro: 3,
  diamante: 4
}

function clean(value) {
  if (value === null || value === undefined) return ''
  return String(value).trim()
}

function normalizeHeader(value) {
  return clean(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
}

function getValue(row, candidates) {
  const keys = Object.keys(row || {})
  for (const candidate of candidates) {
    const normalizedCandidate = normalizeHeader(candidate)
    const found = keys.find(key => normalizeHeader(key) === normalizedCandidate)
    if (found !== undefined) return row[found]
  }
  return ''
}

function isEmptyRow(row) {
  return Object.values(row || {}).every(value => clean(value) === '')
}

function isExcludedExpediente(value) {
  return EXCLUDED_EXPEDIENTES.has(clean(value))
}

function normalizeTipoCliente(value) {
  const raw = normalizeHeader(value)
  if (!raw) return 'SIN CLASIFICAR'
  if (raw.includes('privado')) return 'PRIVADO'
  if (raw.includes('no cobro') || raw.includes('nocobro')) return 'NO COBRO'
  if (raw.includes('institucional')) return 'INSTITUCIONAL'
  if (raw.includes('externo')) return 'EXTERNO'
  if (raw.includes('medico') || raw.includes('medico')) return 'MEDICO'
  if (raw.includes('otros') || raw.includes('otro')) return 'OTROS'
  return clean(value).toUpperCase()
}

function normalizeMoney(value) {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'number') return Number.isFinite(value) ? value : null

  let raw = String(value).trim()
  if (!raw) return null

  raw = raw
    .replace(/\$/g, '')
    .replace(/\s/g, '')
    .replace(/[^\d,.-]/g, '')

  const hasComma = raw.includes(',')
  const hasDot = raw.includes('.')

  if (hasComma && hasDot) {
    raw = raw.replace(/,/g, '')
  } else if (hasComma && !hasDot) {
    raw = raw.replace(',', '.')
  }

  const number = Number(raw)
  return Number.isFinite(number) ? number : null
}

function parseDateValue(value) {
  if (!value) return null

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    if (!date) return null
    const yyyy = String(date.y).padStart(4, '0')
    const mm = String(date.m).padStart(2, '0')
    const dd = String(date.d).padStart(2, '0')
    return `${yyyy}-${mm}-${dd}`
  }

  const raw = clean(value)
  if (!raw) return null

  const iso = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/)
  if (iso) {
    const [, y, m, d] = iso
    return `${y.padStart(4, '0')}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  const local = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/)
  if (local) {
    let [, d, m, y] = local
    if (y.length === 2) y = `20${y}`
    return `${y.padStart(4, '0')}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  const parsed = new Date(raw)
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10)
  return null
}

function isAnulada(value) {
  const raw = normalizeHeader(value)
  return ['si', 's', 'true', '1', 'anulada', 'anulado', 'yes', 'y'].includes(raw)
}

function slug(value) {
  return clean(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/_/g, '-')
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9-]/g, '')
}

function humanLabel(value) {
  if (!value) return 'Sin dato'
  return clean(value)
    .replaceAll('_', ' ')
    .toLowerCase()
    .replace(/\b\w/g, letter => letter.toUpperCase())
}

function formatNumber(value) {
  return new Intl.NumberFormat('es-SV').format(Number(value || 0))
}

function formatMoney(value) {
  return new Intl.NumberFormat('es-SV', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 2
  }).format(Number(value || 0))
}

function formatDate(value) {
  if (!value) return '—'
  const date = new Date(`${String(value).slice(0, 10)}T12:00:00`)
  if (Number.isNaN(date.getTime())) return '—'
  return new Intl.DateTimeFormat('es-SV', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).format(date)
}

function formatDateTime(value) {
  if (!value) return '—'
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return '—'
  return new Intl.DateTimeFormat('es-SV', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date)
}

function percent(value, total) {
  if (!total) return '0.0%'
  return `${((Number(value || 0) / total) * 100).toFixed(1)}%`
}

function chunk(items, size) {
  const chunks = []
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size))
  }
  return chunks
}

function toCsv(rows) {
  if (!rows.length) return ''
  const columns = Object.keys(rows[0])
  const escape = value => `"${String(value ?? '').replaceAll('"', '""')}"`
  return [
    columns.join(','),
    ...rows.map(row => columns.map(column => escape(row[column])).join(','))
  ].join('\n')
}

function downloadCsv(rows, filename) {
  const blob = new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' })
  saveAs(blob, `${filename}.csv`)
}

function downloadXlsx(rows, filename) {
  const worksheet = XLSX.utils.json_to_sheet(rows)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Seguimiento')
  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
  saveAs(new Blob([output]), `${filename}.xlsx`)
}

async function parseFile(file) {
  const name = file.name.toLowerCase()

  if (name.endsWith('.csv')) {
    const text = await file.text()
    return new Promise((resolve, reject) => {
      Papa.parse(text, {
        header: true,
        skipEmptyLines: false,
        delimiter: '',
        transformHeader: header => clean(header),
        complete: result => resolve(result.data || []),
        error: reject
      })
    })
  }

  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
  const sheetName = workbook.SheetNames.includes('BASE') ? 'BASE' : workbook.SheetNames[0]
  const worksheet = workbook.Sheets[sheetName]
  return XLSX.utils.sheet_to_json(worksheet, { defval: '' })
}

async function selectAll(table, columns = '*', pageSize = 1000) {
  let from = 0
  let rows = []

  while (true) {
    const to = from + pageSize - 1
    const { data, error } = await supabase.from(table).select(columns).range(from, to)
    if (error) throw error
    rows = rows.concat(data || [])
    if (!data || data.length < pageSize) break
    from += pageSize
  }

  return rows
}

async function getLatestClientVinculacion() {
  const attempts = ['fecha_vinculacion', 'fecha']
  for (const column of attempts) {
    const { data, error } = await supabase
      .from('clientes_master')
      .select(column)
      .not(column, 'is', null)
      .order(column, { ascending: false })
      .limit(1)

    if (!error && data?.[0]?.[column]) return data[0][column]
  }
  return ''
}

function mapClienteForExport(row) {
  return {
    expediente: row.expediente || '',
    nombre: row.nombre || '',
    telefono: row.telefono || '',
    celular: row.celular || '',
    correo: row.correo || '',
    tipo_cliente: row.tipo_cliente || 'SIN CLASIFICAR',
    ultima_compra: row.ultima_compra || '',
    primera_compra: row.primera_compra || '',
    especialidad_servicio_ultima_compra: row.especialidad_servicio_ultima_compra || '',
    compras_12_meses: row.compras_12m || 0,
    ticket_promedio_12_meses: Number(row.ticket_promedio_12m || 0),
    total_comprado_12_meses: Number(row.total_comprado_12m || 0),
    clasificacion_estrategica: row.clasificacion_estrategica || '',
    etiqueta_visible: row.etiqueta_visible || '',
    estado_relacion: row.estado_relacion || '',
    accion_sugerida: row.accion_sugerida || '',
    empresa: row.sucursal_empresa || row.empresa || '',
    observaciones: row.observaciones || '',
    motivo_inactivo_forzado: row.motivo_inactivo_forzado || ''
  }
}

function transactionDedupKey(record) {
  const uuid = clean(record.uuid_documento).toLowerCase()
  if (uuid) return `uuid:${uuid}`

  const dte = clean(record.dte).toLowerCase()
  const sello = clean(record.sello).toLowerCase()
  if (dte && sello) return `dte:${dte}|${sello}`

  const empresa = clean(record.empresa).toLowerCase()
  const tipo = clean(record.tipo_documento).toLowerCase()
  const serie = clean(record.serie).toLowerCase()
  const numero = clean(record.numero_documento).toLowerCase()
  if (empresa && tipo && serie && numero) return `doc:${empresa}|${tipo}|${serie}|${numero}`

  const noVenta = clean(record.no_venta).toLowerCase()
  const cliente = clean(record.cliente_expediente).toLowerCase()
  if (empresa && noVenta && record.fecha && cliente && record.total !== null) {
    return `old:${empresa}|${noVenta}|${record.fecha}|${cliente}|${Number(record.total).toFixed(2)}`
  }

  return ''
}

export default function App() {
  const [session, setSession] = useState(null)
  const [authEmail, setAuthEmail] = useState('')
  const [authPassword, setAuthPassword] = useState('')
  const [authError, setAuthError] = useState('')

  const [activeTab, setActiveTab] = useState('dashboard')
  const [clientes, setClientes] = useState([])
  const [clientesExport, setClientesExport] = useState([])
  const [comparativoRows, setComparativoRows] = useState([])
  const [historial, setHistorial] = useState([])
  const [ultimaTransaccion, setUltimaTransaccion] = useState('')
  const [ultimaVinculacion, setUltimaVinculacion] = useState('')
  const [filters, setFilters] = useState(DEFAULT_FILTERS)
  const [page, setPage] = useState(1)
  const [loading, setLoading] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session))
    const { data: authListener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession)
    })
    return () => authListener.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) loadInitialData()
  }, [session])

  useEffect(() => {
    setPage(1)
  }, [filters.search, filters.tipoCliente, filters.estado, filters.clasificacion])

  async function signIn(event) {
    event.preventDefault()
    setAuthError('')
    const { error: signInError } = await supabase.auth.signInWithPassword({
      email: authEmail,
      password: authPassword
    })
    if (signInError) setAuthError(signInError.message)
  }

  async function signOut() {
    await supabase.auth.signOut()
  }

  async function loadInitialData() {
    setLoading(true)
    setError('')

    try {
      const [master, exportRows, comparativo, cargas, latestTx, latestVinculacion] = await Promise.all([
        selectAll('clientes_master', 'expediente,nombre,telefono,celular,tipo_cliente').catch(() => []),
        selectAll('v_clientes_operativos_export', '*').catch(() => []),
        selectAll('v_comparativo_anual_clientes', '*').catch(() => []),
        supabase
          .from('cargas_archivos')
          .select('*')
          .order('created_at', { ascending: false })
          .limit(50),
        supabase
          .from('v_transacciones_validas')
          .select('fecha')
          .order('fecha', { ascending: false })
          .limit(1),
        getLatestClientVinculacion()
      ])

      if (cargas.error) throw cargas.error

      const masterRows = (master || []).filter(row => !isExcludedExpediente(row.expediente))
      const tipoByExpediente = new Map(masterRows.map(row => [row.expediente, row.tipo_cliente || 'SIN CLASIFICAR']))
      const mergedExportRows = (exportRows || [])
        .filter(row => !isExcludedExpediente(row.expediente))
        .map(row => ({
        ...row,
        tipo_cliente: row.tipo_cliente || tipoByExpediente.get(row.expediente) || 'SIN CLASIFICAR'
      }))

      setClientes(masterRows)
      setClientesExport(mergedExportRows)
      setHistorial(cargas.data || [])
      setComparativoRows(comparativo || [])
      setUltimaTransaccion(latestTx.data?.[0]?.fecha || '')
      setUltimaVinculacion(latestVinculacion || '')

      const years = [...new Set((comparativo || []).map(row => Number(row.anio)).filter(Boolean))].sort((a, b) => b - a)
      if (years.length && !filters.anioComparativo) {
        setFilters(current => ({ ...current, anioComparativo: String(years[0]) }))
      }
    } catch (loadError) {
      setError(loadError.message || 'No se pudieron cargar los datos.')
    } finally {
      setLoading(false)
    }
  }

  async function registerLoad(payload) {
    const { error: loadError } = await supabase.from('cargas_archivos').insert({
      ...payload,
      uploaded_by: session?.user?.id || null
    })
    if (loadError) throw loadError
  }

  async function handleClientesFile(file) {
    if (!file) return
    setLoading(true)
    setMessage('')
    setError('')

    try {
      const rows = await parseFile(file)
      let invalidRows = 0
      let emptyRows = 0
      const byExpediente = new Map()

      rows.forEach(row => {
        if (isEmptyRow(row)) {
          emptyRows += 1
          return
        }

        const expediente = clean(getValue(row, ['Expediente', 'Numero', 'Número', 'Cliente']))
        const nombre = clean(getValue(row, ['nombre', 'Nombre', 'NCliente', 'Paciente']))
        const telefono = clean(getValue(row, ['telefono', 'Teléfono', 'Telefono']))
        const celular = clean(getValue(row, ['celular', 'Celular']))
        const tipoCliente = normalizeTipoCliente(getValue(row, ['tipo_cliente', 'NTipo de cliente', 'Tipo de cliente', 'TipoCliente']))

        if (!expediente || !nombre) {
          invalidRows += 1
          return
        }

        byExpediente.set(expediente, {
          expediente,
          nombre,
          telefono,
          celular,
          tipo_cliente: tipoCliente,
          origen: 'maestro',
          raw_json: row
        })
      })

      const payload = Array.from(byExpediente.values())
      let updated = 0

      for (const batch of chunk(payload, CLIENT_BATCH_SIZE)) {
        const { error: upsertError } = await supabase
          .from('clientes_master')
          .upsert(batch, { onConflict: 'expediente' })
        if (upsertError) throw upsertError
        updated += batch.length
      }

      await registerLoad({
        tipo_carga: 'clientes',
        nombre_archivo: file.name,
        estado: invalidRows ? 'procesada_con_errores' : 'procesada',
        filas_archivo: rows.length,
        filas_nuevas: 0,
        filas_actualizadas: updated,
        filas_duplicadas: 0,
        filas_invalidas: invalidRows,
        filas_ignoradas: emptyRows,
        ultimo_expediente_detectado: payload.at(-1)?.expediente || null,
        resumen_json: { updated, invalidRows, emptyRows }
      })

      setMessage(`Maestro cargado correctamente: ${formatNumber(updated)} clientes actualizados, ${formatNumber(invalidRows)} inválidos, ${formatNumber(emptyRows)} filas vacías ignoradas.`)
      await loadInitialData()
    } catch (uploadError) {
      setError(uploadError.message || 'Error cargando maestro de clientes.')
    } finally {
      setLoading(false)
    }
  }

  function mapTransaction(row) {
    const empresa = clean(getValue(row, ['Empresa', 'Sucursal']))
    const cliente = clean(getValue(row, ['Cliente', 'Expediente']))
    const nombreCliente = clean(getValue(row, ['NCliente', 'Nombre cliente', 'Nombre']))
    const fecha = parseDateValue(getValue(row, ['Fecha']))
    const total = normalizeMoney(getValue(row, ['Total']))
    const neto = normalizeMoney(getValue(row, ['Neto']))
    const iva = normalizeMoney(getValue(row, ['I.V.A.', 'IVA', 'I.V.A']))

    return {
      no_venta: clean(getValue(row, ['No', 'No.', 'Nº', 'Numero venta', 'Número venta'])),
      empresa,
      tipo_documento: clean(getValue(row, ['Tipo Documento', 'Tipo documento', 'TipoDoc'])),
      serie: clean(getValue(row, ['Serie'])),
      numero_documento: clean(getValue(row, ['Número', 'Numero', 'No Documento', 'Número Documento'])),
      dte: clean(getValue(row, ['DTE', 'Número DTE', 'Numero DTE', 'No DTE'])),
      sello: clean(getValue(row, ['Sello', 'Sello recepción', 'Sello recepcion'])),
      uuid_documento: clean(getValue(row, ['UUID', 'UID'])),
      uuid_valido: Boolean(clean(getValue(row, ['UUID', 'UID']))),
      cliente_expediente: cliente,
      nombre_cliente_transaccion: nombreCliente,
      fecha,
      anulada: isAnulada(getValue(row, ['Anulada', 'Anulado'])),
      neto: neto ?? 0,
      iva: iva ?? 0,
      total,
      especialidad_servicio: clean(getValue(row, ['especialidad_servicio', 'Especialidad', 'Servicio', 'Especialidad o servicio'])),
      observaciones: clean(getValue(row, ['Observaciones', 'Comentario', 'Comentarios'])),
      incluida_analisis: true,
      raw_json: row
    }
  }

  async function insertTransactionBatch(batch) {
    if (!batch.length) return { inserted: 0, duplicated: 0, failed: 0 }

    const { error: insertError } = await supabase.from('transacciones_clientes').insert(batch)

    if (!insertError) return { inserted: batch.length, duplicated: 0, failed: 0 }

    if (batch.length === 1) {
      const message = `${insertError.message || ''} ${insertError.details || ''}`.toLowerCase()
      const duplicated = message.includes('duplicate') || message.includes('unique') || insertError.code === '23505'
      return duplicated ? { inserted: 0, duplicated: 1, failed: 0 } : { inserted: 0, duplicated: 0, failed: 1 }
    }

    const middle = Math.ceil(batch.length / 2)
    const left = await insertTransactionBatch(batch.slice(0, middle))
    const right = await insertTransactionBatch(batch.slice(middle))

    return {
      inserted: left.inserted + right.inserted,
      duplicated: left.duplicated + right.duplicated,
      failed: left.failed + right.failed
    }
  }

  async function handleTransaccionesFile(file) {
    if (!file) return
    setLoading(true)
    setMessage('')
    setError('')

    try {
      const rows = await parseFile(file)
      const transactions = []
      const clientesFromTransactions = new Map()
      const seenKeys = new Set()

      let ignoredRows = 0
      let invalidRows = 0
      let duplicatedInFile = 0

      rows.forEach(row => {
        if (isEmptyRow(row)) {
          ignoredRows += 1
          return
        }

        const record = mapTransaction(row)
        const clienteNormalized = normalizeHeader(record.cliente_expediente)
        const nombreNormalized = normalizeHeader(record.nombre_cliente_transaccion)

        if (record.anulada) {
          ignoredRows += 1
          return
        }

        if (clienteNormalized === 'cliente externo' || nombreNormalized === 'cliente externo') {
          ignoredRows += 1
          return
        }

        const hasUuid = Boolean(record.uuid_documento)
        const hasDteSello = Boolean(record.dte && record.sello)
        const hasDocumentoTradicional = Boolean(record.empresa && record.tipo_documento && record.serie && record.numero_documento)
        const hasFallbackAntiguo = Boolean(record.empresa && record.no_venta && record.fecha && record.cliente_expediente && record.total !== null)
        const valid = record.cliente_expediente && record.fecha && record.total !== null && record.total >= 0 && (
          hasUuid || hasDteSello || hasDocumentoTradicional || hasFallbackAntiguo
        )

        if (!valid) {
          invalidRows += 1
          return
        }

        const key = transactionDedupKey(record)
        if (!key) {
          invalidRows += 1
          return
        }

        if (seenKeys.has(key)) {
          duplicatedInFile += 1
          return
        }

        seenKeys.add(key)
        transactions.push(record)

        if (record.cliente_expediente) {
          clientesFromTransactions.set(record.cliente_expediente, {
            expediente: record.cliente_expediente,
            nombre: record.nombre_cliente_transaccion || record.cliente_expediente,
            tipo_cliente: 'SIN CLASIFICAR',
            origen: 'transaccion'
          })
        }
      })

      let inserted = 0
      let duplicated = duplicatedInFile
      let failed = 0

      for (const batch of chunk(Array.from(clientesFromTransactions.values()), CLIENT_BATCH_SIZE)) {
        const { error: clientError } = await supabase
          .from('clientes_master')
          .upsert(batch, { onConflict: 'expediente', ignoreDuplicates: true })
        if (clientError) throw clientError
      }

      for (const batch of chunk(transactions, TX_BATCH_SIZE)) {
        const result = await insertTransactionBatch(batch)
        inserted += result.inserted
        duplicated += result.duplicated
        failed += result.failed
      }

      await registerLoad({
        tipo_carga: 'transacciones',
        nombre_archivo: file.name,
        estado: invalidRows || failed ? 'procesada_con_errores' : 'procesada',
        filas_archivo: rows.length,
        filas_nuevas: inserted,
        filas_actualizadas: 0,
        filas_duplicadas: duplicated,
        filas_invalidas: invalidRows + failed,
        filas_ignoradas: ignoredRows,
        resumen_json: { inserted, duplicated, duplicatedInFile, invalidRows, failed, ignoredRows }
      })

      setMessage(`Transacciones procesadas: ${formatNumber(inserted)} nuevas, ${formatNumber(duplicated)} duplicadas, ${formatNumber(invalidRows + failed)} inválidas, ${formatNumber(ignoredRows)} ignoradas.`)
      await loadInitialData()
    } catch (uploadError) {
      setError(uploadError.message || 'Error cargando transacciones.')
    } finally {
      setLoading(false)
    }
  }

  const updateInfo = useMemo(() => {
    const latestClientLoad = historial.find(row => row.tipo_carga === 'clientes')
    const latestTxLoad = historial.find(row => row.tipo_carga === 'transacciones')

    return {
      clientesFechaCarga: latestClientLoad?.created_at || '',
      clientesVinculadosHasta: ultimaVinculacion || '',
      ultimoExpediente: latestClientLoad?.ultimo_expediente_detectado || '',
      txFechaCarga: latestTxLoad?.created_at || '',
      txHasta: ultimaTransaccion || ''
    }
  }, [historial, ultimaTransaccion, ultimaVinculacion])

  const filteredClientes = useMemo(() => {
    const search = normalizeHeader(filters.search)

    return clientesExport.filter(row => {
      if (isExcludedExpediente(row.expediente)) return false

      const matchesSearch = !search || [
        row.expediente,
        row.nombre,
        row.telefono,
        row.celular
      ].some(value => normalizeHeader(value).includes(search))

      const matchesTipo = filters.tipoCliente === 'todos' || row.tipo_cliente === filters.tipoCliente
      const matchesEstado = filters.estado === 'todos' || row.estado_relacion === filters.estado
      const matchesClasificacion = filters.clasificacion === 'todos' || row.clasificacion_estrategica === filters.clasificacion

      return matchesSearch && matchesTipo && matchesEstado && matchesClasificacion
    })
  }, [clientesExport, filters])

  const totalPages = Math.max(1, Math.ceil(filteredClientes.length / PAGE_SIZE))
  const paginatedClientes = filteredClientes.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE)

  const tipoOptions = useMemo(() => {
    return [...new Set(clientes.map(row => row.tipo_cliente || 'SIN CLASIFICAR'))].sort()
  }, [clientes])

  const estadoOptions = useMemo(() => {
    return [...new Set(clientesExport.map(row => row.estado_relacion).filter(Boolean))]
      .sort((a, b) => (RELATION_ORDER[a] || 99) - (RELATION_ORDER[b] || 99))
  }, [clientesExport])

  const clasificacionOptions = useMemo(() => {
    return [...new Set(clientesExport.map(row => row.clasificacion_estrategica).filter(Boolean))]
      .sort((a, b) => (STRATEGY_ORDER[a] || 99) - (STRATEGY_ORDER[b] || 99))
  }, [clientesExport])

  const comparativoYears = useMemo(() => {
    return [...new Set(comparativoRows.map(row => Number(row.anio)).filter(Boolean))].sort((a, b) => b - a)
  }, [comparativoRows])

  const filteredComparativo = useMemo(() => {
    const grouped = new Map()

    comparativoRows.forEach(row => {
      const matchesYear = !filters.anioComparativo || String(row.anio) === String(filters.anioComparativo)
      const matchesTipo = filters.tipoClienteComparativo === 'todos' || row.tipo_cliente === filters.tipoClienteComparativo
      if (!matchesYear || !matchesTipo) return

      const key = `${row.anio}-${String(row.mes).padStart(2, '0')}`
      const current = grouped.get(key) || {
        anio: Number(row.anio),
        mes: Number(row.mes),
        mes_nombre: row.mes_nombre,
        tipo_cliente: filters.tipoClienteComparativo === 'todos' ? 'TODOS' : filters.tipoClienteComparativo,
        pacientes_nuevos: 0,
        pacientes_antiguos: 0,
        ingreso_nuevos: 0,
        ingreso_antiguos: 0,
        ingreso_total: 0
      }

      current.pacientes_nuevos += Number(row.pacientes_nuevos || 0)
      current.pacientes_antiguos += Number(row.pacientes_antiguos || 0)
      current.ingreso_nuevos += Number(row.ingreso_nuevos || 0)
      current.ingreso_antiguos += Number(row.ingreso_antiguos || 0)
      current.ingreso_total += Number(row.ingreso_total || 0)

      grouped.set(key, current)
    })

    return [...grouped.values()]
      .map(row => {
        const totalPacientes = row.pacientes_nuevos + row.pacientes_antiguos
        const promedioNuevos = row.pacientes_nuevos ? row.ingreso_nuevos / row.pacientes_nuevos : 0
        const promedioAntiguos = row.pacientes_antiguos ? row.ingreso_antiguos / row.pacientes_antiguos : 0

        return {
          ...row,
          total_pacientes: totalPacientes,
          porcentaje_nuevos: totalPacientes ? row.pacientes_nuevos / totalPacientes : 0,
          promedio_nuevos: promedioNuevos,
          promedio_antiguos: promedioAntiguos,
          diferencia_promedio_n_vs_a: promedioNuevos && promedioAntiguos ? promedioNuevos - promedioAntiguos : 0
        }
      })
      .sort((a, b) => Number(a.mes) - Number(b.mes))
  }, [comparativoRows, filters.anioComparativo, filters.tipoClienteComparativo])

  const kpis = useMemo(() => {
    const total = filteredClientes.length
    const conCompra = filteredClientes.filter(row => row.ultima_compra).length
    const riesgo = filteredClientes.filter(row => ['vigilancia', 'en_riesgo', 'inactivo_reciente'].includes(row.estado_relacion)).length
    const nuevosSinRecompra = filteredClientes.filter(row => row.estado_relacion === 'nuevo_sin_recompra').length
    const embajadores = filteredClientes.filter(row => row.clasificacion_estrategica === 'diamante').length
    const total12m = filteredClientes.reduce((sum, row) => sum + Number(row.total_comprado_12m || 0), 0)
    return { total, conCompra, riesgo, nuevosSinRecompra, embajadores, total12m }
  }, [filteredClientes])

  const estadoDistribution = useMemo(() => {
    const counts = {}
    filteredClientes.forEach(row => {
      const key = row.estado_relacion || 'sin_estado'
      counts[key] = (counts[key] || 0) + 1
    })
    return Object.entries(counts)
      .map(([key, value]) => ({ key, label: humanLabel(key), value }))
      .sort((a, b) => (RELATION_ORDER[a.key] || 99) - (RELATION_ORDER[b.key] || 99))
  }, [filteredClientes])

  const clasificacionDistribution = useMemo(() => {
    const visibleByKey = new Map()
    const counts = {}

    filteredClientes.forEach(row => {
      const key = row.clasificacion_estrategica || 'sin_clasificacion'
      counts[key] = (counts[key] || 0) + 1
      visibleByKey.set(key, row.etiqueta_visible || humanLabel(key))
    })

    return Object.entries(counts)
      .map(([key, value]) => ({ key, label: visibleByKey.get(key), value }))
      .sort((a, b) => (STRATEGY_ORDER[a.key] || 99) - (STRATEGY_ORDER[b.key] || 99))
  }, [filteredClientes])

  const comparativoKpis = useMemo(() => {
    const nuevos = filteredComparativo.reduce((sum, row) => sum + Number(row.pacientes_nuevos || 0), 0)
    const antiguos = filteredComparativo.reduce((sum, row) => sum + Number(row.pacientes_antiguos || 0), 0)
    const ingresoNuevos = filteredComparativo.reduce((sum, row) => sum + Number(row.ingreso_nuevos || 0), 0)
    const ingresoAntiguos = filteredComparativo.reduce((sum, row) => sum + Number(row.ingreso_antiguos || 0), 0)
    const promedioNuevos = nuevos ? ingresoNuevos / nuevos : 0
    const promedioAntiguos = antiguos ? ingresoAntiguos / antiguos : 0
    return { nuevos, antiguos, ingresoNuevos, ingresoAntiguos, promedioNuevos, promedioAntiguos }
  }, [filteredComparativo])

  function updateFilter(key, value) {
    setFilters(current => ({ ...current, [key]: value }))
  }

  function exportFilteredClientes(format) {
    const rows = filteredClientes.map(mapClienteForExport)
    const filename = `Seguimiento clientes ${new Date().toISOString().slice(0, 10)}`
    if (format === 'xlsx') downloadXlsx(rows, filename)
    else downloadCsv(rows, filename)
  }

  function exportComparativo(format) {
    const rows = filteredComparativo.map(row => ({
      anio: row.anio,
      mes: row.mes_nombre,
      tipo_cliente_filtro: row.tipo_cliente,
      pacientes_nuevos: row.pacientes_nuevos,
      pacientes_antiguos: row.pacientes_antiguos,
      total_pacientes: row.total_pacientes,
      porcentaje_nuevos: row.porcentaje_nuevos,
      ingreso_nuevos: row.ingreso_nuevos,
      promedio_nuevos: row.promedio_nuevos,
      ingreso_antiguos: row.ingreso_antiguos,
      promedio_antiguos: row.promedio_antiguos,
      diferencia_promedio_n_vs_a: row.diferencia_promedio_n_vs_a,
      ingreso_total: row.ingreso_total
    }))
    const filename = `Comparativo anual ${filters.anioComparativo || 'todos'}`
    if (format === 'xlsx') downloadXlsx(rows, filename)
    else downloadCsv(rows, filename)
  }

  if (!session) {
    return (
      <div className="login-shell">
        <section className="login-card">
          <img src={logoCcm} alt="CCM" className="login-logo" />
          <h1>CRM Clientes CCM</h1>
          <p>Seguimiento comercial de pacientes, scoring y transacciones.</p>

          {authError && <div className="alert error">{authError}</div>}

          <form onSubmit={signIn} className="login-form">
            <label>
              Correo
              <input value={authEmail} onChange={event => setAuthEmail(event.target.value)} type="email" autoComplete="email" />
            </label>
            <label>
              Contraseña
              <input value={authPassword} onChange={event => setAuthPassword(event.target.value)} type="password" autoComplete="current-password" />
            </label>
            <button className="primary" type="submit">Entrar</button>
          </form>
        </section>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="topbar">
          <div className="brand">
            <div className="logo-box">
              <img src={logoCcm} alt="CCM" />
            </div>
            <div>
              <h1>CRM Clientes CCM</h1>
              <p>Dashboard operativo de pacientes, scoring y seguimiento.</p>
            </div>
          </div>

          <div className="top-actions">
            <div className="update-badge subtle" title="Referencia discreta de actualización de datos">
              {updateInfo.ultimoExpediente ? `Último expediente ${updateInfo.ultimoExpediente}` : 'Último expediente —'}
              {` · Transacciones hasta: ${formatDate(updateInfo.txHasta)}`}
            </div>
            <span className="user-pill">{session.user.email}</span>
            <button className="ghost" onClick={loadInitialData} disabled={loading}>Actualizar</button>
            <button className="secondary" onClick={signOut}>Salir</button>
          </div>
        </div>

        <nav className="tabs">
          {TABS.map(tab => (
            <button
              key={tab.key}
              className={activeTab === tab.key ? 'active' : ''}
              onClick={() => setActiveTab(tab.key)}
            >
              {tab.label}
            </button>
          ))}
        </nav>
      </header>

      <main className="app-main">
        {error && <div className="alert error">{error}</div>}
        {message && <div className="alert success">{message}</div>}
        {loading && <div className="alert info">Procesando / actualizando datos...</div>}

        {['dashboard', 'clientes', 'exportar'].includes(activeTab) && (
          <Filters
            filters={filters}
            updateFilter={updateFilter}
            tipoOptions={tipoOptions}
            estadoOptions={estadoOptions}
            clasificacionOptions={clasificacionOptions}
          />
        )}

        {activeTab === 'dashboard' && (
          <>
            <section className="kpi-grid">
              <KpiCard title="Clientes filtrados" value={formatNumber(kpis.total)} helper="Base operativa según filtros." />
              <KpiCard title="Con compra histórica" value={formatNumber(kpis.conCompra)} helper="Clientes con al menos una transacción válida." />
              <KpiCard title="Riesgo operativo" value={formatNumber(kpis.riesgo)} helper="Vigilancia, en riesgo o inactivo reciente." tone="warning" />
              <KpiCard title="Nuevos sin recompra" value={formatNumber(kpis.nuevosSinRecompra)} helper="Primera compra reciente sin recompra." tone="danger" />
              <KpiCard title="Embajadores" value={formatNumber(kpis.embajadores)} helper="Clientes Diamante según score." tone="success" />
              <KpiCard title="Total 12 meses" value={formatMoney(kpis.total12m)} helper="Compras dentro de ventana móvil 12m." tone="primary" />
            </section>

            <section className="dashboard-grid">
              <DistributionCard title="Estado de relación" subtitle="Distribución relativa sobre clientes filtrados." rows={estadoDistribution} total={filteredClientes.length} type="estado" />
              <StrategyMosaic rows={clasificacionDistribution} total={filteredClientes.length} />
            </section>

            <ClientesTable
              rows={paginatedClientes}
              page={page}
              totalPages={totalPages}
              setPage={setPage}
              totalRows={filteredClientes.length}
              exportFilteredClientes={exportFilteredClientes}
            />
          </>
        )}

        {activeTab === 'clientes' && (
          <ClientesTable
            rows={paginatedClientes}
            page={page}
            totalPages={totalPages}
            setPage={setPage}
            totalRows={filteredClientes.length}
            exportFilteredClientes={exportFilteredClientes}
          />
        )}

        {activeTab === 'comparativo' && (
          <ComparativoTab
            rows={filteredComparativo}
            years={comparativoYears}
            tipoOptions={tipoOptions}
            filters={filters}
            updateFilter={updateFilter}
            kpis={comparativoKpis}
            exportComparativo={exportComparativo}
          />
        )}

        {activeTab === 'exportar' && (
          <section className="card export-card">
            <div>
              <h2>Exportar seguimiento</h2>
              <p>Descarga la tabla de clientes resultante de los filtros activos.</p>
              <p className="muted">{formatNumber(filteredClientes.length)} clientes listos para seguimiento.</p>
            </div>
            <div className="actions">
              <button className="primary" onClick={() => exportFilteredClientes('xlsx')}>Descargar XLSX</button>
              <button className="secondary" onClick={() => exportFilteredClientes('csv')}>Descargar CSV</button>
            </div>
          </section>
        )}

        {activeTab === 'carga' && (
          <section className="upload-grid">
            <UploadCard
              title="Maestro de clientes"
              description="CSV UTF-8 con Expediente, nombre, telefono, celular y tipo_cliente. Actualiza por expediente, no duplica."
              accept=".csv,.xlsx,.xls"
              onFile={handleClientesFile}
            />
            <UploadCard
              title="Transacciones"
              description="CSV/XLSX de facturación. Excluye anuladas y Cliente Externo; deduplica por UUID, DTE+Sello o documento."
              accept=".csv,.xlsx,.xls"
              onFile={handleTransaccionesFile}
            />
          </section>
        )}

        {activeTab === 'historial' && (
          <section className="card">
            <div className="section-head">
              <div>
                <h2>Historial de cargas</h2>
                <p>Últimos archivos procesados y trazabilidad del resultado.</p>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Fecha</th>
                    <th>Tipo</th>
                    <th>Archivo</th>
                    <th>Estado</th>
                    <th>Nuevas</th>
                    <th>Actualizadas</th>
                    <th>Duplicadas</th>
                    <th>Inválidas</th>
                    <th>Ignoradas</th>
                    <th>Último expediente</th>
                  </tr>
                </thead>
                <tbody>
                  {historial.map(row => (
                    <tr key={row.id}>
                      <td>{formatDateTime(row.created_at)}</td>
                      <td><Badge value={row.tipo_carga} type="tipo" /></td>
                      <td>{row.nombre_archivo}</td>
                      <td><Badge value={row.estado} type="estado" /></td>
                      <td>{formatNumber(row.filas_nuevas)}</td>
                      <td>{formatNumber(row.filas_actualizadas)}</td>
                      <td>{formatNumber(row.filas_duplicadas)}</td>
                      <td>{formatNumber(row.filas_invalidas)}</td>
                      <td>{formatNumber(row.filas_ignoradas)}</td>
                      <td>{row.ultimo_expediente_detectado || '—'}</td>
                    </tr>
                  ))}
                  {!historial.length && (
                    <tr>
                      <td colSpan="10" className="empty">No hay cargas registradas.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {activeTab === 'logica' && <LogicTab />}
      </main>
    </div>
  )
}

function Filters({ filters, updateFilter, tipoOptions, estadoOptions, clasificacionOptions }) {
  return (
    <section className="filters card">
      <label>
        Buscar
        <input
          placeholder="Expediente, nombre, teléfono o celular"
          value={filters.search}
          onChange={event => updateFilter('search', event.target.value)}
        />
      </label>

      <label>
        Tipo de cliente
        <select value={filters.tipoCliente} onChange={event => updateFilter('tipoCliente', event.target.value)}>
          <option value="todos">Todos</option>
          {tipoOptions.map(option => <option key={option} value={option}>{option}</option>)}
        </select>
      </label>

      <label>
        Estado de relación
        <select value={filters.estado} onChange={event => updateFilter('estado', event.target.value)}>
          <option value="todos">Todos</option>
          {estadoOptions.map(option => <option key={option} value={option}>{humanLabel(option)}</option>)}
        </select>
      </label>

      <label>
        Clasificación
        <select value={filters.clasificacion} onChange={event => updateFilter('clasificacion', event.target.value)}>
          <option value="todos">Todas</option>
          {clasificacionOptions.map(option => <option key={option} value={option}>{humanLabel(option)}</option>)}
        </select>
      </label>
    </section>
  )
}

function KpiCard({ title, value, helper, tone = 'neutral' }) {
  return (
    <article className={`kpi-card ${tone}`}>
      <span>{title}</span>
      <strong>{value}</strong>
      <p>{helper}</p>
    </article>
  )
}

function DistributionCard({ title, subtitle, rows, total, type }) {
  return (
    <section className="card">
      <div className="section-head">
        <div>
          <h2>{title}</h2>
          <p>{subtitle}</p>
        </div>
      </div>
      <div className="distribution-list">
        {rows.map(row => (
          <div className="distribution-row" key={row.key}>
            <div className="distribution-top">
              <span><Badge value={row.label} type={type} rawKey={row.key} /></span>
              <strong>{formatNumber(row.value)} · {percent(row.value, total)}</strong>
            </div>
            <div className="progress">
              <i style={{ width: percent(row.value, total) }} />
            </div>
          </div>
        ))}
        {!rows.length && <p className="empty">No hay datos para mostrar.</p>}
      </div>
    </section>
  )
}

function StrategyMosaic({ rows, total }) {
  const max = Math.max(...rows.map(row => row.value), 1)

  return (
    <section className="card">
      <div className="section-head">
        <div>
          <h2>Clasificación estratégica</h2>
          <p>Mosaico relativo por valor comercial según recencia, frecuencia y ticket.</p>
        </div>
      </div>
      <div className="strategy-mosaic">
        {rows.map(row => {
          const strength = Math.max(0.3, Number(row.value || 0) / max)
          return (
            <article
              key={row.key}
              className={`mosaic-tile strategy-${slug(row.key)}`}
              style={{ flexGrow: Math.max(1, row.value), minHeight: `${86 + strength * 80}px` }}
            >
              <span>{row.label}</span>
              <strong>{formatNumber(row.value)}</strong>
              <small>{percent(row.value, total)}</small>
            </article>
          )
        })}
        {!rows.length && <p className="empty">No hay datos para mostrar.</p>}
      </div>
    </section>
  )
}

function ClientesTable({ rows, page, totalPages, setPage, totalRows, exportFilteredClientes }) {
  return (
    <section className="card">
      <div className="section-head">
        <div>
          <h2>Clientes</h2>
          <p>{formatNumber(totalRows)} clientes resultantes de los filtros.</p>
        </div>
        <div className="actions">
          <button className="primary" onClick={() => exportFilteredClientes('xlsx')}>Descargar XLSX</button>
          <button className="secondary" onClick={() => exportFilteredClientes('csv')}>Descargar CSV</button>
        </div>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Expediente</th>
              <th>Nombre</th>
              <th>Tipo</th>
              <th>Teléfono</th>
              <th>Celular</th>
              <th>Última compra</th>
              <th>Estado</th>
              <th>Clasificación</th>
              <th>Total 12m</th>
              <th>Acción sugerida</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.expediente}>
                <td><strong>{row.expediente}</strong></td>
                <td>{row.nombre}</td>
                <td><Badge value={row.tipo_cliente || 'SIN CLASIFICAR'} type="tipo" /></td>
                <td>{row.telefono || '—'}</td>
                <td>{row.celular || '—'}</td>
                <td>{formatDate(row.ultima_compra)}</td>
                <td><Badge value={humanLabel(row.estado_relacion)} type="estado" rawKey={row.estado_relacion} /></td>
                <td><Badge value={row.etiqueta_visible || humanLabel(row.clasificacion_estrategica)} type="clasificacion" rawKey={row.clasificacion_estrategica} /></td>
                <td>{formatMoney(row.total_comprado_12m)}</td>
                <td>{row.accion_sugerida || '—'}</td>
              </tr>
            ))}
            {!rows.length && (
              <tr>
                <td colSpan="10" className="empty">No hay registros para los filtros seleccionados.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="pagination">
        <button className="secondary" onClick={() => setPage(current => Math.max(1, current - 1))} disabled={page <= 1}>Anterior</button>
        <span>Página {page} de {totalPages}</span>
        <button className="secondary" onClick={() => setPage(current => Math.min(totalPages, current + 1))} disabled={page >= totalPages}>Siguiente</button>
      </div>
    </section>
  )
}

function ComparativoTab({ rows, years, tipoOptions, filters, updateFilter, kpis, exportComparativo }) {
  return (
    <>
      <section className="filters card">
        <label>
          Año
          <select value={filters.anioComparativo} onChange={event => updateFilter('anioComparativo', event.target.value)}>
            {years.map(year => <option key={year} value={year}>{year}</option>)}
          </select>
        </label>
        <label>
          Tipo de cliente
          <select value={filters.tipoClienteComparativo} onChange={event => updateFilter('tipoClienteComparativo', event.target.value)}>
            <option value="todos">Todos</option>
            {tipoOptions.map(option => <option key={option} value={option}>{option}</option>)}
          </select>
        </label>
      </section>

      <section className="kpi-grid compact">
        <KpiCard title="Pacientes nuevos" value={formatNumber(kpis.nuevos)} helper="Primera compra histórica en el mes." />
        <KpiCard title="Pacientes antiguos" value={formatNumber(kpis.antiguos)} helper="Compraron en el mes, pero ya tenían compras previas." />
        <KpiCard title="Ingreso nuevos" value={formatMoney(kpis.ingresoNuevos)} helper={`Promedio: ${formatMoney(kpis.promedioNuevos)}`} tone="success" />
        <KpiCard title="Ingreso antiguos" value={formatMoney(kpis.ingresoAntiguos)} helper={`Promedio: ${formatMoney(kpis.promedioAntiguos)}`} tone="primary" />
      </section>

      <section className="card">
        <div className="section-head">
          <div>
            <h2>Comparativo anual: nuevos vs antiguos</h2>
            <p>La tabla se consolida por mes según el filtro seleccionado. Promedio = facturación / cantidad de pacientes.</p>
          </div>
          <div className="actions">
            <button className="primary" onClick={() => exportComparativo('xlsx')}>Descargar XLSX</button>
            <button className="secondary" onClick={() => exportComparativo('csv')}>Descargar CSV</button>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Mes</th>
                <th>Nuevos</th>
                <th>Antiguos</th>
                <th>Total pacientes</th>
                <th>% nuevos</th>
                <th>Ingreso nuevos</th>
                <th>Prom. nuevos</th>
                <th>Ingreso antiguos</th>
                <th>Prom. antiguos</th>
                <th>Diferencia prom.</th>
                <th>Ingreso total</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(row => (
                <tr key={`${row.anio}-${row.mes}`}>
                  <td><strong>{row.mes_nombre}</strong></td>
                  <td>{formatNumber(row.pacientes_nuevos)}</td>
                  <td>{formatNumber(row.pacientes_antiguos)}</td>
                  <td>{formatNumber(row.total_pacientes)}</td>
                  <td>{`${(Number(row.porcentaje_nuevos || 0) * 100).toFixed(1)}%`}</td>
                  <td>{formatMoney(row.ingreso_nuevos)}</td>
                  <td>{formatMoney(row.promedio_nuevos)}</td>
                  <td>{formatMoney(row.ingreso_antiguos)}</td>
                  <td>{formatMoney(row.promedio_antiguos)}</td>
                  <td>{formatMoney(row.diferencia_promedio_n_vs_a)}</td>
                  <td>{formatMoney(row.ingreso_total)}</td>
                </tr>
              ))}
              {!rows.length && (
                <tr>
                  <td colSpan="11" className="empty">No hay datos para el filtro seleccionado.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </>
  )
}

function UploadCard({ title, description, accept, onFile }) {
  return (
    <section className="card upload-card">
      <h2>{title}</h2>
      <p>{description}</p>
      <label className="dropzone">
        <input
          type="file"
          accept={accept}
          onChange={event => onFile(event.target.files?.[0])}
        />
        <span>Seleccionar archivo</span>
        <small>CSV UTF-8, XLSX o XLS</small>
      </label>
    </section>
  )
}

function LogicTab() {
  return (
    <section className="logic-grid">
      <article className="card logic-hero">
        <h2>Lógica del reporte</h2>
        <p>
          Este dashboard consolida maestro de clientes y transacciones válidas para clasificar valor,
          estado de relación y oportunidades de seguimiento comercial. La lógica está pensada para
          priorizar gestión, no para copiar hojas de Excel.
        </p>
      </article>

      <article className="card logic-card">
        <h3>1. Cliente y expediente</h3>
        <p>
          <strong>Expediente</strong> es la clave única del paciente/cliente. El maestro se actualiza
          por expediente y no duplica registros.
        </p>
        <p>
          En transacciones, el campo <strong>Cliente</strong> representa el expediente. <strong>NCliente</strong>
          es solo nombre referencial y no se usa como llave.
        </p>
      </article>

      <article className="card logic-card">
        <h3>2. Transacciones válidas</h3>
        <p>
          Se consideran documentos no anulados, con fecha, expediente y total válido. Se excluyen anuladas,
          Cliente Externo y el expediente 3864 por ser cliente varios/no regular.
        </p>
        <p>
          La deduplicación usa UUID cuando existe; si no, usa DTE + sello, documento tradicional o fallback
          documental para registros antiguos.
        </p>
      </article>

      <article className="card logic-card">
        <h3>3. Clasificación estratégica</h3>
        <p>
          Mide valor comercial del cliente con ventana móvil de 12 meses. Usa recencia, frecuencia y ticket
          promedio con pesos configurables.
        </p>
        <p>
          Esta clasificación responde: <strong>¿qué valor estratégico tiene este cliente?</strong>
        </p>
      </article>

      <article className="card logic-card">
        <h3>4. Etiquetas visibles</h3>
        <ul>
          <li><strong>En desarrollo:</strong> clientes de menor valor comercial actual.</li>
          <li><strong>En consolidación:</strong> clientes con potencial, pero aún no plenamente fidelizados.</li>
          <li><strong>Fiel:</strong> clientes de buen valor y comportamiento positivo.</li>
          <li><strong>Embajador:</strong> clientes de mayor valor estratégico.</li>
        </ul>
      </article>

      <article className="card logic-card">
        <h3>5. Estado de relación</h3>
        <p>
          Mide el momento actual del cliente y se mantiene separado del valor estratégico.
          Un cliente puede ser Fiel y estar En riesgo al mismo tiempo.
        </p>
        <p>
          Esta capa responde: <strong>¿qué acción comercial o de seguimiento necesita ahora?</strong>
        </p>
      </article>

      <article className="card logic-card">
        <h3>6. Cliente nuevo</h3>
        <p>
          Un cliente nuevo es aquel cuya primera compra histórica ocurre dentro del período reciente.
          Si compra varias veces en el mismo mes de su primera compra, todo ese ingreso cuenta como nuevo.
        </p>
        <p>
          Si compra después de su mes de primera compra, ya se clasifica como antiguo para efectos comparativos.
        </p>
      </article>
    </section>
  )
}

function Badge({ value, type = 'default', rawKey }) {
  const key = slug(rawKey || value)
  return <span className={`badge badge-${type} badge-${key}`}>{value || '—'}</span>
}
