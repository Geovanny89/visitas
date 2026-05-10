import { useState, useMemo, useEffect, useRef } from 'react'
import * as XLSX from 'xlsx'
import { Search, Download, FileText, MapPin, CheckCircle, Upload, Trash2, AlertCircle, Globe, Sun, Moon } from 'lucide-react'
import { clsx } from 'clsx'
import { twMerge } from 'tailwind-merge'

// Utility for tailwind classes
function cn(...inputs) {
  return twMerge(clsx(inputs))
}

function App() {
  const [businesses, setBusinesses] = useState(() => {
    try {
      const saved = localStorage.getItem('visitas_data')
      if (saved && saved !== 'undefined') {
        const parsed = JSON.parse(saved)
        return Array.isArray(parsed) ? parsed : []
      }
    } catch (e) {
      console.error("Error loading from localStorage", e)
    }
    return []
  })

  const [darkMode, setDarkMode] = useState(() => {
    const savedTheme = localStorage.getItem('theme')
    return savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches)
  })

  const [filterNeighborhood, setFilterNeighborhood] = useState('')
  const [searchTerm, setSearchTerm] = useState('')
  const [error, setError] = useState(null)
  const fileInputRef = useRef(null)

  // Save theme to localStorage
  useEffect(() => {
    localStorage.setItem('theme', darkMode ? 'dark' : 'light')
    if (darkMode) {
      document.documentElement.classList.add('dark')
    } else {
      document.documentElement.classList.remove('dark')
    }
  }, [darkMode])

  // Save to localStorage whenever businesses change
  useEffect(() => {
    if (Array.isArray(businesses)) {
      localStorage.setItem('visitas_data', JSON.stringify(businesses))
    }
  }, [businesses])

  const neighborhoods = useMemo(() => {
    if (!Array.isArray(businesses)) return []
    const uniqueNeighborhoods = Array.from(new Set(businesses.map(b => b?.neighborhood)))
    return uniqueNeighborhoods.filter(Boolean).sort()
  }, [businesses])

  const filteredBusinesses = useMemo(() => {
    if (!Array.isArray(businesses)) return []
    const term = (searchTerm || '').toLowerCase()
    return businesses.filter(b => {
      const matchesNeighborhood = filterNeighborhood === '' || b.neighborhood === filterNeighborhood
      
      const name = String(b?.name || '').toLowerCase()
      const address = String(b?.address || '').toLowerCase()
      
      const matchesSearch = name.includes(term) || address.includes(term)
      return matchesNeighborhood && matchesSearch
    })
  }, [businesses, filterNeighborhood, searchTerm])

  const handleFileUpload = (e) => {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result
        const wb = XLSX.read(bstr, { type: 'binary' })
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        
        // Read as array of arrays to find headers manually
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
        
        if (rows.length === 0) {
          setError("El archivo está vacío.")
          return
        }

        // Find the header row (the one that contains keywords)
        let headerRowIndex = -1
        const keywords = ['barrio', 'name', 'nombre', 'address', 'dirección', 'direccion', 'negocio']
        
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
          const row = rows[i]
          if (Array.isArray(row) && row.some(cell => 
            cell && keywords.includes(String(cell).trim().toLowerCase())
          )) {
            headerRowIndex = i
            break
          }
        }

        // If no header found, assume index 0 but it's risky
        const dataRows = headerRowIndex === -1 ? rows : rows.slice(headerRowIndex + 1)
        const headers = headerRowIndex === -1 ? [] : rows[headerRowIndex].map(h => String(h || '').trim().toLowerCase())

        // Mapping function using header names or positions
        const getVal = (row, searchTerms, posDefault) => {
          // Try by header name first
          if (headers.length > 0) {
            const idx = headers.findIndex(h => searchTerms.includes(h))
            if (idx !== -1 && row[idx] !== undefined) return row[idx]
          }
          // Fallback to default position if valid
          return row[posDefault]
        }

        const mappedData = dataRows
          .filter(row => Array.isArray(row) && row.length > 0 && row.some(cell => cell !== null && cell !== ''))
          .map((row, index) => {
            const name = getVal(row, ['name', 'nombre', 'negocio', 'establecimiento', 'cliente'], 1)
            const address = getVal(row, ['address', 'dirección', 'direccion', 'ubicación', 'ubicacion'], 3)
            const neighborhood = getVal(row, ['barrio', 'neighborhood', 'sector', 'zona'], 0)
            const website = getVal(row, ['website', 'web', 'página', 'pagina'], 2)

            return {
              id: `visita-${Date.now()}-${index}-${Math.random().toString(36).substr(2, 9)}`,
              name: String(name || 'Sin nombre').trim(),
              address: String(address || 'Sin dirección').trim(),
              neighborhood: String(neighborhood || 'Sin barrio').trim(),
              website: String(website || '').trim(),
              visited: false,
              notes: ''
            }
          })
          // Filter out rows that are clearly headers or empty
          .filter(b => b.name !== 'Name' && b.name !== 'Nombre' && b.neighborhood !== 'Barrio')

        if (mappedData.length === 0) {
          setError("No se pudieron encontrar datos válidos. Asegúrate de que las columnas tengan títulos como 'Barrio', 'Name' y 'Address'.")
          return
        }

        setBusinesses(mappedData)
        setError(null)
        if (fileInputRef.current) fileInputRef.current.value = ''
      } catch (err) {
        console.error("Error reading file", err)
        setError("Error al leer el archivo Excel.")
      }
    }
    reader.readAsBinaryString(file)
  }

  const handleToggleVisited = (id) => {
    setBusinesses(prev => prev.map(b => 
      b.id === id ? { ...b, visited: !b.visited } : b
    ))
  }

  const handleUpdateNotes = (id, notes) => {
    setBusinesses(prev => prev.map(b => 
      b.id === id ? { ...b, notes: String(notes) } : b
    ))
  }

  const exportToExcel = () => {
    if (!Array.isArray(businesses) || businesses.length === 0) return

    const dataToExport = businesses.map(b => ({
      'Negocio': b.name,
      'Dirección': b.address,
      'Barrio': b.neighborhood,
      'Website': b.website,
      'Visitado': b.visited ? 'Sí' : 'No',
      'Notas': b.notes,
      'Fecha': new Date().toLocaleDateString()
    }))

    const ws = XLSX.utils.json_to_sheet(dataToExport)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Visitas del Día')
    
    const fileName = `Visitas_${new Date().toISOString().split('T')[0]}.xlsx`
    XLSX.writeFile(wb, fileName)
  }

  const clearData = () => {
    if (window.confirm('¿Estás seguro de que quieres borrar todos los datos? Esto eliminará los negocios cargados y sus notas.')) {
      setBusinesses([])
      setFilterNeighborhood('')
      setSearchTerm('')
      setError(null)
    }
  }

  return (
    <div className={cn(
      "min-h-screen w-full p-4 md:p-8 font-sans transition-colors duration-300",
      darkMode ? "bg-zinc-950 text-zinc-100" : "bg-zinc-50 text-zinc-900"
    )}>
      <div className="max-w-4xl mx-auto">
        <header className="mb-6 flex flex-col gap-4 relative">
          <div className="flex items-start justify-between gap-4">
            <div className="flex-1 min-w-0">
              <h1 className={cn(
                "text-2xl md:text-3xl font-bold transition-colors truncate",
                darkMode ? "text-blue-400" : "text-blue-600"
              )}>Gestión de Visitas</h1>
              <p className={cn(
                "text-xs md:text-sm mt-1",
                darkMode ? "text-zinc-400" : "text-zinc-500"
              )}>Carga tu ruta y registra tus visitas</p>
            </div>
            <button 
              onClick={() => setDarkMode(!darkMode)}
              className={cn(
                "shrink-0 p-3 rounded-2xl transition-all active:scale-95 shadow-xl flex items-center justify-center",
                darkMode ? "bg-zinc-800 text-yellow-400 shadow-zinc-950/50" : "bg-white text-zinc-600 shadow-zinc-200 border border-zinc-200"
              )}
              title={darkMode ? "Cambiar a modo claro" : "Cambiar a modo oscuro"}
            >
              {darkMode ? <Sun size={22} /> : <Moon size={22} />}
            </button>
          </div>
          
          <div className="flex flex-wrap gap-2 w-full">
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2.5 rounded-xl transition-all shadow-lg shadow-blue-900/20 active:scale-95 text-sm font-semibold"
            >
              <Upload size={18} />
              Cargar
            </button>
            <input 
              type="file" 
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx, .xls, .csv"
              className="hidden"
            />
            {businesses.length > 0 && (
              <>
                <button 
                  onClick={exportToExcel}
                  className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2.5 rounded-xl transition-all shadow-lg shadow-emerald-900/20 active:scale-95 text-sm font-semibold"
                >
                  <Download size={18} />
                  Exportar
                </button>
                <button 
                  onClick={clearData}
                  className={cn(
                    "flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl transition-all active:scale-95 text-sm font-semibold border",
                    darkMode 
                      ? "bg-zinc-800 hover:bg-zinc-700 text-zinc-300 border-zinc-700" 
                      : "bg-white hover:bg-zinc-50 text-zinc-600 border-zinc-200"
                  )}
                >
                  <Trash2 size={18} />
                  Borrar
                </button>
              </>
            )}
          </div>
        </header>

        {error && (
          <div className={cn(
            "mb-6 px-4 py-3 rounded-xl flex items-center gap-3 animate-in fade-in slide-in-from-top-2 border",
            darkMode 
              ? "bg-red-900/20 border-red-500/50 text-red-200" 
              : "bg-red-50 border-red-200 text-red-700"
          )}>
            <AlertCircle size={20} className="shrink-0" />
            <p className="text-sm">{error}</p>
          </div>
        )}

        {businesses.length > 0 ? (
          <>
            <div className={cn(
              "rounded-2xl shadow-xl p-4 mb-6 border transition-colors",
              darkMode ? "bg-zinc-900 border-zinc-800" : "bg-white border-zinc-200"
            )}>
              <div className="flex flex-col gap-3">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-500" size={18} />
                  <input 
                    type="text"
                    placeholder="Buscar negocio o dirección..."
                    className={cn(
                      "w-full pl-10 pr-4 py-3 rounded-xl transition-all text-sm outline-none border",
                      darkMode 
                        ? "bg-zinc-950 border-zinc-800 text-zinc-100 placeholder:text-zinc-600 focus:ring-blue-500/50" 
                        : "bg-zinc-50 border-zinc-200 text-zinc-900 placeholder:text-zinc-400 focus:ring-blue-500/20"
                    )}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                </div>
                <div className="relative">
                  <MapPin className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-500" size={18} />
                  <select 
                    className={cn(
                      "w-full pl-10 pr-10 py-3 rounded-xl transition-all text-sm outline-none border appearance-none",
                      darkMode 
                        ? "bg-zinc-950 border-zinc-800 text-zinc-100 focus:ring-blue-500/50" 
                        : "bg-zinc-50 border-zinc-200 text-zinc-900 focus:ring-blue-500/20"
                    )}
                    value={filterNeighborhood}
                    onChange={(e) => setFilterNeighborhood(e.target.value)}
                  >
                    <option value="">Todos los barrios ({neighborhoods.length})</option>
                    {neighborhoods.map(n => (
                      <option key={n} value={n}>{n}</option>
                    ))}
                  </select>
                  <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-zinc-500">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-4 h-4">
                      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5" />
                    </svg>
                  </div>
                </div>
              </div>
            </div>

            <div className="space-y-3">
              <div className="flex justify-between items-center px-1 mb-2">
                <p className="text-xs text-zinc-500 font-bold uppercase tracking-wider">
                  {filteredBusinesses.length} negocios encontrados
                </p>
                <p className="text-xs text-emerald-500 font-bold uppercase tracking-wider">
                  {businesses.filter(b => b.visited).length} visitados
                </p>
              </div>
              
              {filteredBusinesses.map(business => (
                <div 
                  key={business.id}
                  className={cn(
                    "rounded-2xl shadow-sm border p-4 transition-all active:scale-[0.98]",
                    darkMode 
                      ? business.visited ? "bg-emerald-500/5 border-emerald-500/30" : "bg-zinc-900 border-zinc-800"
                      : business.visited ? "bg-emerald-50 border-emerald-200" : "bg-white border-zinc-200"
                  )}
                >
                  <div className="flex flex-col gap-3">
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className={cn(
                            "text-base font-bold truncate max-w-[200px] md:max-w-none",
                            business.visited 
                              ? darkMode ? "text-emerald-400" : "text-emerald-700" 
                              : darkMode ? "text-zinc-100" : "text-zinc-900"
                          )}>
                            {business.name}
                          </h3>
                          {business.visited && <CheckCircle size={16} className="text-emerald-500 shrink-0" />}
                        </div>
                        <div className="flex flex-col gap-1 mt-1">
                          <div className={cn(
                            "flex items-center gap-1.5 text-sm",
                            darkMode ? "text-zinc-500" : "text-zinc-600"
                          )}>
                            <MapPin size={14} className="shrink-0" />
                            <span className="truncate">{business.address}</span>
                          </div>
                          <div className="flex items-center gap-2 mt-1">
                            <span className={cn(
                              "px-2 py-0.5 rounded-lg text-[10px] font-bold uppercase tracking-tight border",
                              darkMode 
                                ? "bg-zinc-800 text-zinc-400 border-zinc-700" 
                                : "bg-zinc-100 text-zinc-500 border-zinc-200"
                            )}>
                              {business.neighborhood}
                            </span>
                            {business.website && (
                              <a 
                                href={business.website.startsWith('http') ? business.website : `https://${business.website}`} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="text-blue-500 hover:text-blue-400 flex items-center gap-1 text-[10px] font-bold uppercase"
                              >
                                <Globe size={12} /> Web
                              </a>
                            )}
                          </div>
                        </div>
                      </div>
                      <button 
                        onClick={() => handleToggleVisited(business.id)}
                        className={cn(
                          "shrink-0 px-4 py-2 rounded-xl text-xs font-bold transition-all shadow-sm",
                          business.visited 
                            ? "bg-emerald-600 text-white" 
                            : darkMode 
                              ? "bg-zinc-800 text-zinc-300 border-zinc-700" 
                              : "bg-zinc-100 text-zinc-600 border-zinc-200"
                        )}
                      >
                        {business.visited ? "Visto" : "Marcar"}
                      </button>
                    </div>
                    
                    <div className="relative">
                      <FileText className={cn(
                        "absolute left-3 top-3",
                        darkMode ? "text-zinc-600" : "text-zinc-400"
                      )} size={16} />
                      <textarea 
                        placeholder="Notas de la visita..."
                        className={cn(
                          "w-full pl-10 pr-4 py-2.5 rounded-xl outline-none min-h-[70px] text-sm resize-none border transition-all",
                          darkMode 
                            ? "bg-zinc-950 border-zinc-800 text-zinc-300 placeholder:text-zinc-700 focus:ring-blue-500/30" 
                            : "bg-zinc-50 border-zinc-200 text-zinc-700 placeholder:text-zinc-400 focus:ring-blue-500/10"
                        )}
                        value={business.notes}
                        onChange={(e) => handleUpdateNotes(business.id, e.target.value)}
                      />
                    </div>
                  </div>
                </div>
              ))}
              
              {filteredBusinesses.length === 0 && (
                <div className={cn(
                  "text-center py-12 rounded-2xl border border-dashed transition-colors",
                  darkMode ? "bg-zinc-900/50 border-zinc-800" : "bg-zinc-100/50 border-zinc-300"
                )}>
                  <p className="text-zinc-500 text-sm">No se encontraron negocios.</p>
                </div>
              )}
            </div>
          </>
        ) : (
          <div className={cn(
            "flex flex-col items-center justify-center py-16 rounded-3xl border-2 border-dashed shadow-2xl text-center px-6 transition-colors",
            darkMode ? "bg-zinc-900 border-zinc-800" : "bg-white border-zinc-200 shadow-zinc-200"
          )}>
            <div className={cn(
              "p-5 rounded-2xl mb-5 border",
              darkMode ? "bg-blue-500/10 border-blue-500/20" : "bg-blue-50 border-blue-100"
            )}>
              <Upload className="text-blue-500" size={40} />
            </div>
            <h2 className={cn(
              "text-xl font-bold mb-2",
              darkMode ? "text-zinc-100" : "text-zinc-800"
            )}>Comienza tu ruta</h2>
            <p className="text-zinc-500 max-w-xs mb-8 text-sm">
              Sube tu archivo Excel con <span className={cn("font-bold", darkMode ? "text-zinc-300" : "text-zinc-700")}>Barrio, Name y Address</span> para empezar.
            </p>
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3.5 rounded-2xl font-bold transition-all shadow-xl shadow-blue-900/40 active:scale-95 w-full max-w-[240px]"
            >
              Seleccionar archivo
            </button>
          </div>
        )}
      </div>
    </div>
  )
}

export default App
