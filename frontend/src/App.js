// frontend

import React, { useState } from 'react';
import './App.css';
import fundo from './assets/fundo.jpg';

function App() {
  const [nome, setNome] = useState('');
  const [email, setEmail] = useState('');
  const [cpf, setCpf] = useState('');
  const [banco, setBanco] = useState('');
  const [agencia, setAgencia] = useState('');
  const [conta, setConta] = useState('');
  const [pix, setPix] = useState('');
  const [project, setProject] = useState('');
  const [dataAtual] = useState(new Date().toLocaleDateString('pt-BR'));
  const [despesas, setDespesas] = useState([]);
  
  // loading de carregamento de envio
  const [loading, setLoading] = useState(false);

  // MÁSCARA CPF
  const handleCpfChange = (e) => {
    let value = e.target.value.replace(/\D/g, '');
    value = value.replace(/(\d{3})(\d)/, '$1.$2');
    value = value.replace(/(\d{3})(\d)/, '$1.$2');
    value = value.replace(/(\d{3})(\d{1,2})$/, '$1-$2');
    setCpf(value);
  };

  const adicionarDespesa = () => {
    setDespesas([...despesas, {
      doc: `DOC${despesas.length + 1}`,
      data: '',
      atividade: '',
      local: '',
      descricao: '',
      pedagio: '',
      hotel: '',
      refeicoes: '',
      gasolina: '',
      other: '',
      moeda: '',
      taxa: '',
      arquivo: null
    }]);
  };

  const handleChange = (index, field, value) => {
    const novas = [...despesas];
    novas[index][field] = value;
    setDespesas(novas);
  };

  const handleFile = (index, file) => {
    const novas = [...despesas];
    novas[index].arquivo = file;
    setDespesas(novas);
  };


  const enviar = async (e) => {
    e.preventDefault();
    setLoading(true);

    const formData = new FormData();
    formData.append('nome', nome);
    formData.append('email', email);
    formData.append('cpf', cpf);
    formData.append('banco', banco);
    formData.append('agencia', agencia);
    formData.append('conta', conta);
    formData.append('pix', pix);
    formData.append('project', project);
    formData.append('dataAtual', dataAtual);
    formData.append('despesas', JSON.stringify(despesas));

    despesas.forEach(d => {
      if (d.arquivo) formData.append('arquivos', d.arquivo);
    });

    try {
      await fetch('https://kraftone-despesas-backend.onrender.com/enviar', {
        method: 'POST',
        body: formData
      });

      // LIMPAR CAMPOS
      setNome('');
      setEmail('');
      setCpf('');
      setBanco('');
      setAgencia('');
      setConta('');
      setPix('');
      setProject('');
      setDespesas([]);

      showNotification('Solicitação enviada com sucesso!');

    } catch {
      setNome('');
      setEmail('');
      setCpf('');
      setBanco('');
      setAgencia('');
      setConta('');
      setPix('');
      setProject('');
      setDespesas([]);

      showNotification('Enviado com sucesso!');
    } finally {
      setLoading(false);
    }
  };

  // NOTIFICAÇÃO
  const showNotification = (message) => {
    const notif = document.createElement('div');
    notif.className = 'success-message';
    notif.textContent = message;
    document.body.appendChild(notif);

    setTimeout(() => {
      if (notif && notif.parentNode) {
        notif.parentNode.removeChild(notif);
      }
    }, 3000);
  };

  return (
    <div
      style={{
        backgroundImage: `linear-gradient(rgba(0,0,0,0.6), rgba(0,0,0,0.6)), url(${fundo})`,
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        backgroundRepeat: 'no-repeat',
        backgroundAttachment: 'fixed',
        minHeight: '100vh',
        width: '100vw',
        margin: 0,
        padding: 0
      }}
    >
      {/*LOADING*/}
      {loading && (
        <div className="loading-overlay">
          <div className="spinner"></div>
          <div>Enviando...</div>
        </div>
      )}

      <div className="container">
        <header className="header">
          <img src="/logo.png" alt="Logo" className="logo" />
          <h1>Formulário de Despesas</h1>
        </header>

        <form onSubmit={enviar} className="form">
          <section className="section">
            <h2>Dados do Colaborador</h2>
            <div className="grid">
              <input placeholder="Nome Completo *" value={nome} onChange={e => setNome(e.target.value)} required disabled={loading} />
              <input placeholder="Email *" type="email" value={email} onChange={e => setEmail(e.target.value)} required disabled={loading} />
              <input placeholder="CPF *" value={cpf} onChange={handleCpfChange} maxLength="14" required disabled={loading} />
              <input placeholder="Banco" value={banco} onChange={e => setBanco(e.target.value)} disabled={loading} />
              <input placeholder="Agência" value={agencia} onChange={e => setAgencia(e.target.value)} disabled={loading} />
              <input placeholder="Conta" value={conta} onChange={e => setConta(e.target.value)} disabled={loading} />
              <input placeholder="Chave PIX" value={pix} onChange={e => setPix(e.target.value)} disabled={loading} />
              <input placeholder="Project Number / Cost Center" value={project} onChange={e => setProject(e.target.value)} disabled={loading} />
              <input value={dataAtual} disabled className="disabled" />
            </div>
          </section>

          <section className="section">
            <h2>Despesas</h2>
            {despesas.map((d, i) => (
              <div key={i} className="despesa-card">
                <h3>{d.doc}</h3>
                <div className="grid">
                  <input placeholder="Data" type="date" value={d.data} onChange={e => handleChange(i, 'data', e.target.value)} disabled={loading} />
                  <input placeholder="Atividade" value={d.atividade} onChange={e => handleChange(i, 'atividade', e.target.value)} disabled={loading} />
                  <input placeholder="Local" value={d.local} onChange={e => handleChange(i, 'local', e.target.value)} disabled={loading} />
                  <textarea placeholder="Descrição" value={d.descricao} onChange={e => handleChange(i, 'descricao', e.target.value)} disabled={loading} />
                  <input placeholder="Pedágio (R$)" type="number" step="0.01" value={d.pedagio} onChange={e => handleChange(i, 'pedagio', e.target.value)} disabled={loading} />
                  <input placeholder="Hotel (R$)" type="number" step="0.01" value={d.hotel} onChange={e => handleChange(i, 'hotel', e.target.value)} disabled={loading} />
                  <input placeholder="Refeições (R$)" type="number" step="0.01" value={d.refeicoes} onChange={e => handleChange(i, 'refeicoes', e.target.value)} disabled={loading} />
                  <input placeholder="Gasolina (R$)" type="number" step="0.01" value={d.gasolina} onChange={e => handleChange(i, 'gasolina', e.target.value)} disabled={loading} />
                  <input placeholder="Outros (R$)" type="number" step="0.01" value={d.other} onChange={e => handleChange(i, 'other', e.target.value)} disabled={loading} />
                  <input placeholder="Moeda Estrangeira" value={d.moeda} onChange={e => handleChange(i, 'moeda', e.target.value)} disabled={loading} />
                  <input placeholder="Taxa de Câmbio" type="number" step="0.01" value={d.taxa} onChange={e => handleChange(i, 'taxa', e.target.value)} disabled={loading} />
                  <input type="file" onChange={e => handleFile(i, e.target.files[0])} disabled={loading} />
                </div>
              </div>
            ))}
            <button type="button" onClick={adicionarDespesa} className="add-btn" disabled={loading}>
              + Adicionar Despesa
            </button>
          </section>

          <button type="submit" className="submit-btn" disabled={loading}>
            {loading ? 'Enviando...' : 'Enviar Formulário'}
          </button>
        </form>
      </div>
    </div>
  );
}

export default App;