import { useEffect, useMemo, useState } from 'react'
import { supabase } from './lib/supabase'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import logoCcm from './assets/logo-ccm.png'

const CLIENT_COLUMNS = ['expediente', 'nombre', 'telefono', 'celular', 'tipo_cliente']
const CLIENT_BATCH_SIZE = 500
const TX_BATCH_SIZE = 250

function clean(value) {
  return String(value ?? '').trim()
}

function normalizeKey(value) {
  return clean(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[.\-\/]/g, ' ')
    .replace(/\s+/g, '_')
}

function normalizeText(value) {
  return clean(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
}

function isRowEmpty(row) {
  if (!row || typeof row !== 'object') return true
  return Object.values(row).every(value => clean(value) === '')
}

function getValue(row, candidates) {
  const normalized = {}

  Object.entries(row || {}).forEach(([key, value]) => {
    normalized[normalizeKey(key)] = value
  })

  for (const candidate of candidates) {
    const key = normalizeKey(candidate)
    if (Object.prototype.hasOwnProperty.call(normalized, key)) {
      return normalized[key]
    }
  }

  return ''
}

function normalizeTipoCliente(value) {
  const raw = normalizeText(value)

  if (!raw) return 'SIN CLASIFICAR'
  if (raw.includes('PRIVADO')) return 'PRIVADO'
  if (raw.includes('NO COBRO') || raw.includes('NO_COBRO')) return 'NO COBRO'
  if (raw.includes('INSTITUCIONAL')) return 'INSTITUCIONAL'
  if (raw.includes('EXTERNO')) return 'EXTERNO'
  if (raw.includes('MEDICO') || raw.includes('MÉDICO')) return 'MEDICO'
  if (raw.includes('OTROS') || raw.includes('OTRO')) return 'OTROS'

  return raw
}

function parseDate(value) {
  if (value === null || value === undefined || value === '') return null

  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    if (!date) return null
    return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`
  }

  const raw = clean(value)
  if (!raw) return null

  const iso = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/)
  if (iso) {
    const [, y, m, d] = iso
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  const latin = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/)
  if (latin) {
    let [, d, m, y] = latin
    if (y.length === 2) y = `20${y}`
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  const parsed = new Date(raw)
  if (Number.isNaN(parsed.getTime())) return null
  return parsed.toISOString().slice(0, 10)
}

function parseMoney(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0

  let raw = clean(value)
  if (!raw) return 0

  raw = raw.replace(/[^\d,.-]/g, '')

  const hasComma = raw.includes(',')
  const hasDot = raw.includes('.')

  if (hasComma && hasDot) {
    const lastComma = raw.lastIndexOf(',')
    const lastDot = raw.lastIndexOf('.')
    if (lastComma > lastDot) {
      raw = raw.replace(/\./g, '').replace(',', '.')
    } else {
      raw = raw.replace(/,/g, '')
    }
  } else if (hasComma && !hasDot) {
    raw = raw.replace(',', '.')
  }

  const parsed = Number(raw)
  return Number.isFinite(parsed) ? parsed : 0
}

function isAnulada(value) {
  const raw = normalizeText(value)
  return ['SI', 'S', 'TRUE', 'VERDADERO', 'ANULADA', 'ANULADO', '1', 'YES'].includes(raw)
}

function isTotalizationRow(row) {
  const values = Object.values(row || {}).map(normalizeText).join(' ')
  return values.includes('TOTAL GENERAL') || values === 'TOTAL' || values.includes('TOTALIZACION')
}

function chunk(array, size) {
  const chunks = []
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size))
  }
  return chunks
}

async function readFile(file) {
  const name = file.name.toLowerCase()

  if (name.endsWith('.csv')) {
    const text = await file.text()

    return new Promise((resolve, reject) => {
      Papa.parse(text, {
        header: true,
        skipEmptyLines: 'greedy',
        transformHeader: header => clean(header),
        complete: result => resolve(result.data || []),
        error: reject
      })
    })
  }

  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  return XLSX.utils.sheet_to_json(sheet, { defval: '' })
}

function Login() {
  const [email, setEmail] = useState('admin@grupo-ccm.com')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState('')

  async function handleLogin(event) {
    event.preventDefault()
    setMessage('Validando acceso...')

    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) setMessage(error.message)
  }

  return (
    <main className="login-page">
      <form className="login-card" onSubmit={handleLogin}>
        <div className="login-logo">
          <img src={logoCcm} alt="CCM" />
        </div>
        <h1>CRM Clientes CCM</h1>
        <p>Acceso seguro al módulo de seguimiento</p>

        <label>Email</label>
        <input value={email} onChange={event => setEmail(event.target.value)} />

        <label>Contraseña</label>
        <input type="password" value={password} onChange={event => setPassword(event.target.value)} />

        <button type="submit">Entrar</button>

        {message && <div className="error">{message}</div>}
      </form>
    </main>
  )
}

export default function App() {
  const [session, setSession] = useState(null)
  const [tab, setTab] = useState('carga')
  const [message, setMessage] = useState('')
  const [clientes, setClientes] = useState([])
  const [cargas, setCargas] = useState([])
  const [loading, setLoading] = useState(false)

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session))

    const { data: authListener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession)
    })

    return () => authListener.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadDashboard()
      loadCargas()
    }
  }, [session])

  const stats = useMemo(() => {
    return {
      clientes: clientes.length,
      enRiesgo: clientes.filter(row => row.estado_relacion === 'en_riesgo').length,
      nuevos: clientes.filter(row => String(row.estado_relacion || '').startsWith('nuevo')).length,
      embajadores: clientes.filter(row => row.clasificacion_estrategica === 'diamante').length
    }
  }, [clientes])

  async function loadDashboard() {
    const { data, error } = await supabase
      .from('v_clientes_operativos_export')
      .select('*')
      .limit(1000)

    if (error) setMessage(error.message)
    else setClientes(data || [])
  }

  async function loadCargas() {
    const { data, error } = await supabase
      .from('cargas_archivos')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(30)

    if (!error) setCargas(data || [])
  }

  async function uploadClientes(file) {
    if (!file) return

    setLoading(true)

    try {
      setMessage('Leyendo maestro de clientes...')
      const rawRows = await readFile(file)

      let invalidas = 0
      let ignoradas = 0
      let ultimoExpediente = ''

      const mapped = []

      for (const row of rawRows) {
        if (isRowEmpty(row)) {
          ignoradas += 1
          continue
        }

        const expediente = clean(getValue(row, ['Expediente', 'expediente', 'Número', 'Numero', 'Cliente']))
        const nombre = clean(getValue(row, ['nombre', 'Nombre', 'NCliente', 'Cliente Nombre']))
        const tipoCliente = normalizeTipoCliente(getValue(row, [
          'tipo_cliente',
          'Tipo cliente',
          'Tipo de cliente',
          'NTipo de cliente',
          'NTipo cliente',
          'NTipoCliente'
        ]))

        if (!expediente || !nombre) {
          invalidas += 1
          continue
        }

        ultimoExpediente = expediente

        mapped.push({
          expediente,
          nombre,
          telefono: clean(getValue(row, ['telefono', 'teléfono', 'Telefono', 'Teléfono'])),
          celular: clean(getValue(row, ['celular', 'Celular'])),
          tipo_cliente: tipoCliente,
          correo: clean(getValue(row, ['correo Electronico', 'correo electronico', 'correo', 'email', 'Correo'])),
          comentarios: clean(getValue(row, ['comentarios', 'comentario', 'Comentarios'])),
          fecha_vinculacion: parseDate(getValue(row, ['fecha', 'Fecha', 'fecha_vinculacion'])),
          origen: 'maestro',
          raw_json: row
        })
      }

      let cargados = 0

      for (const [index, group] of chunk(mapped, CLIENT_BATCH_SIZE).entries()) {
        setMessage(`Cargando maestro de clientes... lote ${index + 1} de ${Math.ceil(mapped.length / CLIENT_BATCH_SIZE)}`)

        const { error } = await supabase
          .from('clientes_master')
          .upsert(group, { onConflict: 'expediente' })

        if (error) throw error
        cargados += group.length
      }

      await supabase.from('cargas_archivos').insert({
        tipo_carga: 'clientes',
        nombre_archivo: file.name,
        filas_archivo: rawRows.length,
        filas_nuevas: cargados,
        filas_actualizadas: cargados,
        filas_invalidas: invalidas,
        filas_ignoradas: ignoradas,
        ultimo_expediente_detectado: ultimoExpediente,
        resumen_json: {
          campos_oficiales: CLIENT_COLUMNS,
          modo: 'upsert_por_expediente'
        },
        uploaded_by: session?.user?.id || null
      })

      setMessage(`Maestro cargado correctamente: ${cargados} clientes actualizados, ${invalidas} inválidos, ${ignoradas} filas vacías ignoradas.`)

      await loadDashboard()
      await loadCargas()
    } catch (error) {
      setMessage(`Error cargando clientes: ${error.message}`)
    } finally {
      setLoading(false)
    }
  }

  async function ensureClientesFromTransacciones(items) {
    const unique = new Map()

    for (const item of items) {
      if (!item.cliente_expediente) continue
      if (!unique.has(item.cliente_expediente)) {
        unique.set(item.cliente_expediente, {
          expediente: item.cliente_expediente,
          nombre: item.nombre_cliente_transaccion || item.cliente_expediente,
          origen: 'transaccion',
          raw_json: { creado_desde_transaccion: true }
        })
      }
    }

    const clientesMinimos = Array.from(unique.values())

    for (const group of chunk(clientesMinimos, CLIENT_BATCH_SIZE)) {
      const { error } = await supabase
        .from('clientes_master')
        .upsert(group, {
          onConflict: 'expediente',
          ignoreDuplicates: true
        })

      if (error) throw error
    }
  }

  async function insertTransaccionesBatch(group) {
    const { error } = await supabase
      .from('transacciones_clientes')
      .insert(group)

    if (!error) return { nuevas: group.length, duplicadas: 0, invalidas: 0 }

    let nuevas = 0
    let duplicadas = 0
    let invalidas = 0

    for (const item of group) {
      const { error: rowError } = await supabase
        .from('transacciones_clientes')
        .insert(item)

      if (!rowError) nuevas += 1
      else if (rowError.code === '23505') duplicadas += 1
      else invalidas += 1
    }

    return { nuevas, duplicadas, invalidas }
  }

  async function uploadTransacciones(file) {
    if (!file) return

    setLoading(true)

    try {
      setMessage('Leyendo archivo de transacciones...')
      const rawRows = await readFile(file)

      let invalidas = 0
      let ignoradas = 0

      const mapped = []

      for (const row of rawRows) {
        if (isRowEmpty(row)) {
          ignoradas += 1
          continue
        }

        if (isTotalizationRow(row)) {
          ignoradas += 1
          continue
        }

        const expediente = clean(getValue(row, ['Cliente', 'Expediente', 'Número', 'Numero']))
        const nombre = clean(getValue(row, ['NCliente', 'nombre', 'Nombre']))
        const fecha = parseDate(getValue(row, ['Fecha', 'fecha']))
        const anulada = isAnulada(getValue(row, ['Anulada', 'anulada']))
        const total = parseMoney(getValue(row, ['Total', 'total']))
        const neto = parseMoney(getValue(row, ['Neto', 'neto']))
        const iva = parseMoney(getValue(row, ['I.V.A.', 'IVA', 'iva']))
        const sucursal = clean(getValue(row, ['Sucursal', 'sucursal'])) || 'SIN SUCURSAL'
        const tipoDocumento = clean(getValue(row, ['Tipo Documento', 'TipoDocumento', 'tipo_documento'])) || 'SIN TIPO'
        const serie = clean(getValue(row, ['Serie', 'serie'])) || 'SIN SERIE'
        const numeroDocumento = clean(getValue(row, ['Número Documento', 'Numero Documento', 'Número', 'Numero', 'No Documento'])) || ''
        const uuid = clean(getValue(row, ['UUID', 'uuid']))
        const observaciones = clean(getValue(row, ['Observaciones', 'observaciones']))
        const especialidadServicio = clean(getValue(row, [
          'especialidad_servicio',
          'Especialidad',
          'Servicio',
          'Especialidad o servicio',
          'Descripcion',
          'Descripción'
        ]))

        if (anulada || normalizeText(nombre).includes('CLIENTE EXTERNO')) {
          ignoradas += 1
          continue
        }

        if (!expediente || !fecha || !numeroDocumento || total < 0) {
          invalidas += 1
          continue
        }

        mapped.push({
          uuid_documento: uuid || null,
          uuid_valido: Boolean(uuid),
          sucursal,
          tipo_documento: tipoDocumento,
          serie,
          numero_documento: numeroDocumento,
          cliente_expediente: expediente,
          nombre_cliente_transaccion: nombre || expediente,
          fecha,
          especialidad_servicio: especialidadServicio || null,
          neto,
          iva,
          total,
          observaciones,
          anulada: false,
          incluida_analisis: true,
          raw_json: row
        })
      }

      await ensureClientesFromTransacciones(mapped)

      let nuevas = 0
      let duplicadas = 0
      let invalidasInsert = 0

      const groups = chunk(mapped, TX_BATCH_SIZE)

      for (const [index, group] of groups.entries()) {
        setMessage(`Cargando transacciones... lote ${index + 1} de ${groups.length}`)

        const result = await insertTransaccionesBatch(group)
        nuevas += result.nuevas
        duplicadas += result.duplicadas
        invalidasInsert += result.invalidas
      }

      await supabase.from('cargas_archivos').insert({
        tipo_carga: 'transacciones',
        nombre_archivo: file.name,
        fecha_inicio_dato: mapped.length ? mapped.map(row => row.fecha).sort()[0] : null,
        fecha_fin_dato: mapped.length ? mapped.map(row => row.fecha).sort().at(-1) : null,
        filas_archivo: rawRows.length,
        filas_nuevas: nuevas,
        filas_duplicadas: duplicadas,
        filas_invalidas: invalidas + invalidasInsert,
        filas_ignoradas: ignoradas,
        resumen_json: {
          modo: 'insert_batch_con_fallback_por_fila',
          deduplicacion: 'UUID o Sucursal+Tipo Documento+Serie+Número'
        },
        uploaded_by: session?.user?.id || null
      })

      setMessage(`Transacciones cargadas: ${nuevas} nuevas, ${duplicadas} duplicadas, ${invalidas + invalidasInsert} inválidas, ${ignoradas} ignoradas.`)

      await loadDashboard()
      await loadCargas()
    } catch (error) {
      setMessage(`Error cargando transacciones: ${error.message}`)
    } finally {
      setLoading(false)
    }
  }

  function exportar(formato) {
    const rows = clientes.map(row => ({
      expediente: row.expediente,
      nombre: row.nombre,
      telefono: row.telefono,
      celular: row.celular,
      correo: row.correo,
      ultima_compra: row.ultima_compra,
      primera_compra: row.primera_compra,
      compras_12m: row.compras_12m,
      ticket_promedio_12m: row.ticket_promedio_12m,
      total_comprado_12m: row.total_comprado_12m,
      clasificacion_estrategica: row.clasificacion_estrategica,
      etiqueta_visible: row.etiqueta_visible,
      estado_relacion: row.estado_relacion,
      accion_sugerida: row.accion_sugerida,
      sucursal_empresa: row.sucursal_empresa,
      observaciones: row.observaciones,
      motivo_inactivo_forzado: row.motivo_inactivo_forzado
    }))

    const worksheet = XLSX.utils.json_to_sheet(rows)

    if (formato === 'csv') {
      const csv = XLSX.utils.sheet_to_csv(worksheet)
      saveAs(new Blob([csv], { type: 'text/csv;charset=utf-8' }), 'Seguimiento clientes.csv')
      return
    }

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Seguimiento clientes')
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    saveAs(new Blob([buffer]), 'Seguimiento clientes.xlsx')
  }

  if (!session) return <Login />

  return (
    <main className="app">
      <header>
        <div className="brand">
          <div className="brand-logo">
            <img src={logoCcm} alt="CCM" />
          </div>
          <div>
            <h1>CRM Clientes CCM</h1>
            <p>Carga, scoring y seguimiento comercial</p>
          </div>
        </div>

        <button className="secondary" onClick={() => supabase.auth.signOut()}>Salir</button>
      </header>

      <nav>
        {['carga', 'dashboard', 'exportar', 'historial'].map(item => (
          <button key={item} className={tab === item ? 'active' : ''} onClick={() => setTab(item)}>
            {item}
          </button>
        ))}
      </nav>

      {message && <div className="message">{message}</div>}

      {tab === 'carga' && (
        <section className="grid">
          <div className="card">
            <h2>Maestro de clientes</h2>
            <p>CSV UTF-8 recomendado. Campos oficiales: Expediente, nombre, telefono, celular, tipo_cliente.</p>
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              disabled={loading}
              onChange={event => uploadClientes(event.target.files?.[0])}
            />
            <small>No duplica clientes: actualiza por Expediente.</small>
          </div>

          <div className="card">
            <h2>Transacciones</h2>
            <p>CSV o XLSX. Excluye anuladas, Cliente Externo y filas vacías.</p>
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              disabled={loading}
              onChange={event => uploadTransacciones(event.target.files?.[0])}
            />
            <small>No duplica documentos: usa UUID o fallback documental.</small>
          </div>
        </section>
      )}

      {tab === 'dashboard' && (
        <section>
          <div className="stats">
            <div><span>Clientes</span><strong>{stats.clientes}</strong></div>
            <div><span>En riesgo</span><strong>{stats.enRiesgo}</strong></div>
            <div><span>Nuevos</span><strong>{stats.nuevos}</strong></div>
            <div><span>Embajadores</span><strong>{stats.embajadores}</strong></div>
          </div>

          <Table
            rows={clientes}
            columns={[
              'expediente',
              'nombre',
              'ultima_compra',
              'clasificacion_estrategica',
              'etiqueta_visible',
              'estado_relacion',
              'accion_sugerida'
            ]}
          />
        </section>
      )}

      {tab === 'exportar' && (
        <section className="card">
          <h2>Exportar Seguimiento clientes</h2>
          <p>Descarga la base operativa actual para seguimiento comercial.</p>
          <button onClick={() => exportar('xlsx')}>Descargar Excel</button>
          <button onClick={() => exportar('csv')}>Descargar CSV</button>
        </section>
      )}

      {tab === 'historial' && (
        <Table
          rows={cargas}
          columns={[
            'created_at',
            'tipo_carga',
            'nombre_archivo',
            'filas_archivo',
            'filas_nuevas',
            'filas_actualizadas',
            'filas_duplicadas',
            'filas_invalidas',
            'filas_ignoradas'
          ]}
        />
      )}
    </main>
  )
}

function Table({ rows, columns }) {
  return (
    <div className="card table-wrap">
      <table>
        <thead>
          <tr>{columns.map(column => <th key={column}>{column}</th>)}</tr>
        </thead>
        <tbody>
          {rows.slice(0, 100).map((row, index) => (
            <tr key={index}>
              {columns.map(column => <td key={column}>{String(row[column] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
