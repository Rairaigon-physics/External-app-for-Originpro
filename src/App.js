import React, { useState } from 'react';
import axios from 'axios';
import './App.css';

// --- Component: Embedded Scientific Logo ---
const LabLogo = () => (
  <svg className="app-logo" viewBox="0 0 512 512" fill="none" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="bg_grad" x1="0" y1="0" x2="512" y2="512" gradientUnits="userSpaceOnUse">
        <stop offset="0" stopColor="#1e1b4b"/>
        <stop offset="1" stopColor="#312e81"/>
      </linearGradient>
      <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
        <feGaussianBlur stdDeviation="6" result="coloredBlur"/>
        <feMerge><feMergeNode in="coloredBlur"/><feMergeNode in="SourceGraphic"/></feMerge>
      </filter>
    </defs>
    <path d="M256 24 L478 144 V368 L256 488 L34 368 V144 Z" fill="url(#bg_grad)" stroke="#4338ca" strokeWidth="8"/>
    <g opacity="0.2" stroke="white" strokeWidth="2" strokeDasharray="5 5">
      <path d="M100 144 H412"/> <path d="M100 204 H412"/> <path d="M100 264 H412"/>
      <path d="M100 324 H412"/> <path d="M156 100 V412"/> <path d="M256 100 V412"/> <path d="M356 100 V412"/>
    </g>
    <g stroke="white" strokeWidth="6" strokeLinecap="round">
      <path d="M130 400 L130 130"/> <path d="M130 400 L400 400"/>
    </g>
    <g filter="url(#glow)">
        <path d="M150 370 C 200 350, 250 200, 380 160" stroke="#f472b6" strokeWidth="10" fill="none" strokeLinecap="round"/>
        <path d="M380 190 C 280 220, 220 320, 150 340" stroke="#60a5fa" strokeWidth="10" fill="none" strokeLinecap="round" strokeDasharray="20 15"/>
    </g>
    <rect x="140" y="360" width="20" height="20" fill="white"/> <rect x="250" y="265" width="20" height="20" fill="white"/> <rect x="370" y="150" width="20" height="20" fill="white"/>
    <circle cx="150" cy="340" r="10" fill="#93c5fd" stroke="white" strokeWidth="3"/> <circle cx="260" cy="260" r="10" fill="#93c5fd" stroke="white" strokeWidth="3"/> <circle cx="380" cy="190" r="10" fill="#93c5fd" stroke="white" strokeWidth="3"/>
  </svg>
);

const LoadingSpinner = () => (
  <svg className="spinner" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" opacity="0.3" />
    <path d="M12 2a10 10 0 0 1 10 10" stroke="currentColor" strokeWidth="4" strokeLinecap="round" />
  </svg>
);

// --- Helper: Shared Upload Logic (Updated for Options) ---
const uploadData = async (endpoint, fileMap, textData, config, setLoading) => {
    if (!Object.values(fileMap).every(f => f)) {
        alert("Please select all required files.");
        return;
    }
    setLoading(true);

    const formData = new FormData();
    Object.entries(fileMap).forEach(([key, file]) => formData.append(key, file));
    Object.entries(textData).forEach(([key, value]) => formData.append(key, value));
    
    // Pass the Toggle Options to Backend
    formData.append('createPPT', config.createPPT);
    formData.append('saveProject', config.saveProject);

    const firstFile = Object.values(fileMap)[0];
    if (firstFile) {
        const lastModifiedDate = new Date(firstFile.lastModified).toISOString();
        formData.append('lastModified', lastModifiedDate);
    }

    try {
        const response = await axios.post(`http://localhost:5000/${endpoint}`, formData, {
            headers: { 'Content-Type': 'multipart/form-data' },
            timeout: 0 
        });
        console.log(response.data);
        alert(response.data.message || "Operation Successful");
    } catch (error) {
        console.error('Error uploading files:', error);
        alert("Error: " + (error.response?.data?.error || "Failed. Check console."));
    } finally {
        setLoading(false);
    }
};

// --- Component: Toggle Switch ---
// --- Component: Toggle Switch (Fixed) ---
const ToggleOption = ({ label, checked, onChange }) => (
    <div className="toggle-wrapper">
        <label className="switch">
            <input 
                type="checkbox" 
                checked={checked} 
                onChange={(e) => onChange(e.target.checked)} 
            />
            <span className="slider"></span>
        </label>
        {/* Allow clicking the text label to toggle it too */}
        <span 
            className="toggle-label" 
            onClick={() => onChange(!checked)}
            style={{cursor: 'pointer'}}
        >
            {label}
        </span>
    </div>
);

// --- Component: Generic Single File Form ---
const ExperimentForm = ({ title, endpoint, textInputs = [], config }) => {
    const [file, setFile] = useState(null);
    const [values, setValues] = useState({});
    const [isLoading, setIsLoading] = useState(false);

    const handleTextChange = (key, val) => setValues(prev => ({ ...prev, [key]: val }));
    // Pass config to uploadData
    const handleSubmit = () => uploadData(endpoint, { datafile: file }, values, config, setIsLoading);

    return (
        <div className="section card-hover">
            {title && <h2>{title}</h2>}
            <div className="file-input-wrapper">
                <label className="input-label">Data File</label>
                <input type="file" onChange={(e) => setFile(e.target.files[0])} className="file-input" />
            </div>
            {textInputs.map(({ key, placeholder }) => (
                <input key={key} type="text" placeholder={placeholder} value={values[key] || ''}
                    onChange={(e) => handleTextChange(key, e.target.value)} className="text-input" />
            ))}
            <button onClick={handleSubmit} className="btn" disabled={isLoading}>
                {isLoading ? <><LoadingSpinner /> Processing...</> : 'Plot Data'}
            </button>
        </div>
    );
};

// --- Component: Dewar Dual Form ---
const DewarDualForm = ({ config }) => {
    const [cooling, setCooling] = useState(null);
    const [warming, setWarming] = useState(null);
    const [pressure, setPressure] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    
    // Pass config to uploadData
    const handleSubmit = () => uploadData('dewar', { cooling, warming }, { pressure }, config, setIsLoading);

    return (
        <div className="section dewar-section card-hover">
            <div className="dewar-header">
                <h1>Dewar Resistance</h1>
                <p>Dual channel cooling & warming analysis</p>
            </div>
            <div className="dual-upload-grid">
                <div className="upload-col"><span className="badge badge-blue">Cooling Data</span><input type="file" onChange={(e) => setCooling(e.target.files[0])} className="file-input" /></div>
                <div className="upload-col"><span className="badge badge-red">Warming Data</span><input type="file" onChange={(e) => setWarming(e.target.files[0])} className="file-input" /></div>
            </div>
            <input type="text" placeholder="Enter pressure in (GPa)" value={pressure} onChange={(e) => setPressure(e.target.value)} className="text-input main-input" />
            <button onClick={handleSubmit} className="btn btn-large" disabled={isLoading}>{isLoading ? <><LoadingSpinner /> Plotting...</> : 'Plot Dual Data'}</button>
        </div>
    );
};

// --- Main App Component ---
function App() {
    const [showMenu, setShowMenu] = useState(false);
    
    // Global Settings State
    const [config, setConfig] = useState({
        createPPT: true,
        saveProject: true
    });

    const togglePPT = (val) => setConfig(prev => ({ ...prev, createPPT: val }));
    const toggleSave = (val) => setConfig(prev => ({ ...prev, saveProject: val }));

    return (
        <div className="App">
            <header className="app-header">
                <LabLogo />
                <h1 className="app-title">Lab Automation Dashboard</h1>
                <p className="app-subtitle">OriginPro Data Plotting & Export Interface</p>
            </header>

            {/* OPTIONS PANEL */}
            <div style={{textAlign: 'center'}}>
                <div className="options-panel">
                    <ToggleOption label="Create PowerPoint" checked={config.createPPT} onChange={togglePPT} />
                    <ToggleOption label="Save Origin Project" checked={config.saveProject} onChange={toggleSave} />
                </div>
            </div>

            <main className="app-content">
                <DewarDualForm config={config} />

                <div className="toggle-container">
                    <button onClick={() => setShowMenu(!showMenu)} className="menu-btn">
                        {showMenu ? 'Close Merged Data Option' : 'Open Merged Data Option'}
                    </button>
                    <div className={`merged-section-wrapper ${showMenu ? 'open' : ''}`}>
                        <div className="merged-section">
                             <ExperimentForm title="Dewar Merged Data" endpoint="dewar_strip" 
                                textInputs={[{ key: 'pressure', placeholder: 'Enter pressure (GPa)' }]} config={config} />
                        </div>
                    </div>
                </div>

                <div className="divider"></div>

                 <div className="wide-section">
                    <ExperimentForm title="Current Effect Analysis" endpoint="current_effect" 
                        textInputs={[{ key: 'pressure', placeholder: 'Enter pressure (GPa)' }]} config={config} />
                 </div>

                <div className="divider"></div>
                <h2 className="category-title">PPMS Measurements</h2>

                <div className="grid-section">
                    <ExperimentForm title="Resistance" endpoint="ppms" 
                        textInputs={[{ key: 'pressure', placeholder: 'Enter pressure (GPa)' }]} config={config} />
                    <ExperimentForm title="Magnetic Field" endpoint="ppms_magnetic" 
                        textInputs={[{ key: 'pressure', placeholder: 'Enter pressure (GPa)' }]} config={config} />
                    <ExperimentForm title="Heat Capacity" endpoint="ppms_heat_capacity" 
                        textInputs={[{ key: 'mass_heat_cap', placeholder: 'Enter mass (mg)' }]} config={config} />
                </div>
                
                <div className="wide-section" style={{marginTop: '25px'}}>
                    <ExperimentForm title="Heat Capacity (Cooling/Warming)" endpoint="ppms_heat_capacity_cw" 
                        textInputs={[{ key: 'mass', placeholder: 'Enter mass (mg)' }]} config={config} />
                </div>

                <div className="divider"></div>
                <h2 className="category-title">MPMS Measurements</h2>

                <div className="grid-section">
                    <ExperimentForm title="Moment vs Temp" endpoint="mpms" 
                        textInputs={[{ key: 'magnetic_moment', placeholder: 'Enter magnetic moment (Oe)' }]} config={config} />
                    <ExperimentForm title="Magnetic Field" endpoint="mpms_magnetic" 
                        textInputs={[{ key: 'mass', placeholder: 'Enter Mass (mg)' }]} config={config} />
                </div>
                
                <div className="wide-section">
                    <ExperimentForm title="AC Susceptibility" endpoint="mpms_ac" textInputs={[
                            { key: 'mass_ac', placeholder: 'Enter Mass (mg)' },
                            { key: 'MF_dc', placeholder: 'Enter DC Field (Oe)' },
                            { key: 'MF_ac', placeholder: 'Enter AC Field (Oe)' },
                        ]} config={config} />
                </div>
            </main>
        </div>
    );
}

export default App;