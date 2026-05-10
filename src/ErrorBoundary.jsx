import React from 'react'

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props)
    this.state = { hasError: false, error: null }
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error }
  }

  componentDidCatch(error, errorInfo) {
    console.error("ErrorBoundary caught an error", error, errorInfo)
  }

  handleReset = () => {
    localStorage.removeItem('visitas_data')
    window.location.reload()
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4 text-center font-sans">
          <div className="bg-white p-8 rounded-2xl shadow-xl max-w-md border border-red-100">
            <div className="w-16 h-16 bg-red-100 text-red-600 rounded-full flex items-center justify-center mx-auto mb-4">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-10 h-10">
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" />
              </svg>
            </div>
            <h1 className="text-2xl font-bold text-gray-900 mb-2">¡Ups! Algo salió mal</h1>
            <p className="text-gray-600 mb-6">
              Hubo un error al procesar los datos. Esto suele pasar si el archivo Excel tenía un formato inesperado.
            </p>
            <div className="flex flex-col gap-3">
              <button
                onClick={this.handleReset}
                className="bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-6 rounded-xl transition-all shadow-lg active:scale-95"
              >
                Limpiar datos y Reiniciar
              </button>
              <button
                onClick={() => window.location.reload()}
                className="text-gray-500 hover:text-gray-700 font-medium text-sm"
              >
                Solo recargar página
              </button>
            </div>
          </div>
        </div>
      )
    }

    return this.props.children
  }
}

export default ErrorBoundary
