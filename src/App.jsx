import { useEffect, useState } from 'react'
import { supabase } from './lib/supabase'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'

function clean(v) {
  return String(v ?? '').trim()
}

function normalizeKey(v) {
  return clean(v)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\./g, '')
    .replace(/\s+/g, '_')
}

function getValue(row, candidates) {
  const normalized = {}
  Object.entries(row).forEach(([key, value]) => {
    normalized[normalizeKey(key)] = value
  })

  for (const candidate of candidates) {
    const key = normalizeKey(candidate)
    if (Object.prototype.hasOwnProperty.call(normalized, key)) return normalized[key]
  }

  return ''
}

function parseDate(value) {
  const s = clean(value)
  if (!s) return null
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10)

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/)
  if (m) {
    const year = m[3].length === 2 ? `20${m[3]}` : m[3]
    return `${year}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`
  }

  return null
}

function parseMoney(value) {
  if (typeof value === 'number') return value
  let s = clean(value).replace(/\$/g, '').replace(/\s/g, '')
  if (s.includes(',') && s.lastIndexOf(',') > s.lastIndexOf('.')) {
    s = s.replace(/\./g, '').replace(',', '.')
  } else {
    s = s.replace(/,/g, '')
  }
  const n = Number(s)
  return Number.isFinite(n) ? n : 0
}

async function readFile(file) {
  const name = file.name.toLowerCase()

  if (name.endsWith('.csv')) {
    const text = await file.text()
    return new Promise((resolve, reject) => {
      Papa.parse(text, {
        header: true,
        skipEmptyLines: true,
        complete: result => resolve(result.data || []),
        error: reject
      })
    })
  }

  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  return XLSX.utils.sheet_to_json(sheet, { defval: '' })
}

function chunk(items, size) {
  const result = []
  for (let i = 0; i < items.length; i += size) result.push(items.slice(i, i + size))
  return result
}

function Login() {
  const [email, setEmail] = useState('admin@grupo-ccm.com')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState('')

  async function submit(e) {
    e.preventDefault()
    setMessage('')
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) setMessage(error.message)
  }

  return (
    <main className="login">
      <form onSubmit={submit}>
        <h1>CRM Clientes CCM</h1>
        <p>Acceso interno</p>
        <label>Email</label>
        <input value={email} onChange={e => setEmail(e.target.value)} />
        <label>Contraseña</label>
        <input type="password" value={password} onChange={e => setPassword(e.target.value)} />
        <button>Entrar</button>
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

  async function loadDashboard() {
    const { data, error } = await supabase
      .from('v_clientes_operativos_export')
      .select('*')
      .limit(1000)

    if (error) setMessage(error.message)
    else setClientes(data || [])
  }

  async function loadCargas() {
    const { data } = await supabase
      .from('cargas_archivos')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(20)

    setCargas(data || [])
  }

  async function uploadClientes(file) {
    try {
      if (!file) return
      setMessage('Leyendo maestro de clientes...')

      const rows = await readFile(file)
      let invalidas = 0
      let ultimoExpediente = ''

      const mapped = rows.map(row => {
        const expediente = clean(getValue(row, ['Número', 'Numero', 'expediente']))
        const nombre = clean(getValue(row, ['nombre', 'Nombre']))

        if (!expediente || !nombre) {
          invalidas += 1
          return null
        }

        ultimoExpediente = expediente

        return {
          expediente,
          nombre,
          telefono: clean(getValue(row, ['telefono', 'teléfono'])),
          celular: clean(getValue(row, ['celular'])),
          correo: clean(getValue(row, ['correo Electronico', 'correo electronico', 'correo', 'email'])),
          comentarios: clean(getValue(row, ['comentarios', 'comentario'])),
          fecha_vinculacion: parseDate(getValue(row, ['fecha', 'Fecha'])),
          origen: 'maestro',
          raw_json: row
        }
      }).filter(Boolean)

      for (const group of chunk(mapped, 300)) {
        const { error } = await supabase
          .from('clientes_master')
          .upsert(group, { onConflict: 'expediente' })

        if (error) throw error
      }

      await supabase.from('cargas_archivos').insert({
        tipo_carga: 'clientes',
        nombre_archivo: file.name,
        filas_archivo: rows.length,
        filas_nuevas: mapped.length,
        filas_invalidas: invalidas,
        ultimo_expediente_detectado: ultimoExpediente,
        uploaded_by: session?.user?.id || null
      })

      setMessage(`Clientes cargados: ${mapped.length}. Inválidos: ${invalidas}.`)
      await loadDashboard()
      await loadCargas()
    } catch (error) {
      setMessage(error.message)
    }
  }

  async function uploadTransacciones(file) {
    try {
      if (!file) return
      setMessage('Leyendo transacciones...')

      const rows = await readFile(file)
      let nuevas = 0
      let duplicadas = 0
      let invalidas = 0
      let ignoradas = 0

      for (const row of rows) {
        const expediente = clean(getValue(row, ['Cliente', 'cliente']))
        const nombre = clean(getValue(row, ['NCliente', 'N Cliente', 'nombre cliente']))
        const fecha = parseDate(getValue(row, ['Fecha', 'fecha']))
        const anulada = ['si', 'sí', 's', 'true', '1', 'x'].includes(clean(getValue(row, ['Anulada', 'anulada'])).toLowerCase())

        if (anulada || nombre.toLowerCase().includes('cliente externo')) {
          ignoradas += 1
          continue
        }

        if (!expediente || !fecha) {
          invalidas += 1
          continue
        }

        await supabase
          .from('clientes_master')
          .upsert(
            { expediente, nombre: nombre || expediente, origen: 'transaccion' },
            { onConflict: 'expediente', ignoreDuplicates: true }
          )

        const item = {
          uuid_documento: clean(getValue(row, ['UUID', 'uuid'])) || null,
          uuid_valido: Boolean(clean(getValue(row, ['UUID', 'uuid']))),
          sucursal: clean(getValue(row, ['Sucursal', 'sucursal'])) || 'SIN_SUCURSAL',
          tipo_documento: clean(getValue(row, ['Tipo Documento', 'tipo documento'])) || 'SIN_TIPO',
          serie: clean(getValue(row, ['Serie', 'serie'])) || 'SIN_SERIE',
          numero_documento: clean(getValue(row, ['Número', 'Numero', 'numero'])) || '',
          cliente_expediente: expediente,
          nombre_cliente_transaccion: nombre,
          fecha,
          especialidad_servicio: clean(getValue(row, ['Especialidad', 'Servicio', 'especialidad', 'servicio'])) || null,
          neto: parseMoney(getValue(row, ['Neto', 'neto'])),
          iva: parseMoney(getValue(row, ['I.V.A.', 'IVA', 'iva'])),
          total: parseMoney(getValue(row, ['Total', 'total'])),
          observaciones: clean(getValue(row, ['Observaciones', 'observaciones'])),
          anulada: false,
          incluida_analisis: true,
          raw_json: row
        }

        const { error } = await supabase.from('transacciones_clientes').insert(item)

        if (error?.code === '23505') duplicadas += 1
        else if (error) invalidas += 1
        else nuevas += 1
      }

      await supabase.from('cargas_archivos').insert({
        tipo_carga: 'transacciones',
        nombre_archivo: file.name,
        filas_archivo: rows.length,
        filas_nuevas: nuevas,
        filas_duplicadas: duplicadas,
        filas_invalidas: invalidas,
        filas_ignoradas: ignoradas,
        uploaded_by: session?.user?.id || null
      })

      setMessage(`Transacciones: ${nuevas} nuevas, ${duplicadas} duplicadas, ${invalidas} inválidas, ${ignoradas} ignoradas.`)
      await loadDashboard()
      await loadCargas()
    } catch (error) {
      setMessage(error.message)
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
        <div>
          <h1>CRM Clientes CCM</h1>
          <p>Carga, scoring y seguimiento comercial</p>
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
            <p>Subir CSV de pacientes/clientes.</p>
            <input type="file" accept=".csv" onChange={e => uploadClientes(e.target.files?.[0])} />
          </div>

          <div className="card">
            <h2>Transacciones</h2>
            <p>Subir CSV o XLSX de facturación/transacciones.</p>
            <input type="file" accept=".csv,.xlsx,.xls" onChange={e => uploadTransacciones(e.target.files?.[0])} />
          </div>
        </section>
      )}

      {tab === 'dashboard' && (
        <section>
          <div className="stats">
            <div><span>Clientes</span><strong>{clientes.length}</strong></div>
            <div><span>En riesgo</span><strong>{clientes.filter(x => x.estado_relacion === 'en_riesgo').length}</strong></div>
            <div><span>Nuevos</span><strong>{clientes.filter(x => String(x.estado_relacion || '').startsWith('nuevo')).length}</strong></div>
            <div><span>Embajadores</span><strong>{clientes.filter(x => x.clasificacion_estrategica === 'diamante').length}</strong></div>
          </div>

          <Table rows={clientes} columns={['expediente', 'nombre', 'ultima_compra', 'especialidad_servicio_ultima_compra', 'etiqueta_visible', 'estado_relacion', 'accion_sugerida']} />
        </section>
      )}

      {tab === 'exportar' && (
        <section className="card">
          <h2>Exportar Seguimiento clientes</h2>
          <button onClick={() => exportar('xlsx')}>Descargar Excel</button>
          <button onClick={() => exportar('csv')}>Descargar CSV</button>
        </section>
      )}

      {tab === 'historial' && (
        <Table rows={cargas} columns={['created_at', 'tipo_carga', 'nombre_archivo', 'filas_nuevas', 'filas_duplicadas', 'filas_invalidas', 'filas_ignoradas']} />
      )}
    </main>
  )
}

function Table({ rows, columns }) {
  return (
    <div className="card table-wrap">
      <table>
        <thead>
          <tr>{columns.map(col => <th key={col}>{col}</th>)}</tr>
        </thead>
        <tbody>
          {rows.slice(0, 100).map((row, index) => (
            <tr key={index}>
              {columns.map(col => <td key={col}>{String(row[col] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
