import pkg from '@whiskeysockets/baileys';
import qrcode from 'qrcode-terminal';
import fs from 'fs';
import xlsx from 'xlsx';

const { makeWASocket, makeInMemoryStore, useMultiFileAuthState, DisconnectReason } = pkg;

// Criar um armazenamento em memória
const store = makeInMemoryStore({});

// Carregar o estado de autenticação
const { state, saveCreds } = await useMultiFileAuthState('./auth');

// Carregar os contatos e perguntas a partir dos arquivos Excel
const contatosWorkbook = xlsx.readFile('./contatos.xlsx');
const contatosSheet = contatosWorkbook.Sheets[contatosWorkbook.SheetNames[0]];
const contatos = xlsx.utils.sheet_to_json(contatosSheet);

const perguntasWorkbook = xlsx.readFile('./pesquisa.xlsx');
const perguntasSheet = perguntasWorkbook.Sheets['Perguntas'];
const perguntas = xlsx.utils.sheet_to_json(perguntasSheet);

const respostasWorkbook = xlsx.readFile('./pesquisa.xlsx');
const respostasSheet = respostasWorkbook.Sheets['Respostas'];
const respostas = xlsx.utils.sheet_to_json(respostasSheet, { header: 1 });

// Garantir que a aba "Respostas" tenha cabeçalhos
if (!respostas.length) {
    respostas.push(['Pergunta', ...contatos.map(contato => contato.Nome)]);
}

// Controlar o progresso dos contatos
const progresso = contatos.reduce((acc, contato) => {
    acc[contato.Numero] = 0; // Começar com a primeira pergunta
    return acc;
}, {});

// Função para conectar ao WhatsApp
async function conectarWhatsApp() {
    const socket = makeWASocket({
        auth: state,
    });

    store.bind(socket.ev);

    socket.ev.on('creds.update', saveCreds);

    socket.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr) {
            console.log('⚠️ Conexão perdida. Por favor, escaneie o novo código QR:');
            qrcode.generate(qr, { small: true });
        }

        if (connection === 'open') {
            console.log('✅ Conectado ao WhatsApp!');
            for (const contato of contatos) {
                enviarPergunta(socket, contato.Numero);
            }
        }

        if (connection === 'close') {
            const deveReconectar = (lastDisconnect?.error)?.output?.statusCode !== DisconnectReason.loggedOut;
            if (deveReconectar) {
                console.log('⚠️ Conexão fechada. Tentando reconectar...');
                await conectarWhatsApp();
            } else {
                console.log('❌ Desconectado. Por favor, exclua a pasta "auth" e reinicie o bot para se reconectar.');
            }
        }
    });

    // Escutar por mensagens recebidas
    socket.ev.on('messages.upsert', async (msg) => {
        const mensagem = msg.messages[0];
        if (!mensagem.key.fromMe) {
            tratarResposta(socket, mensagem);
        }
    });
}

// Função para enviar uma pergunta a um contato
async function enviarPergunta(socket, numeroContato) {
    const indicePerguntaAtual = progresso[numeroContato];

    if (indicePerguntaAtual >= perguntas.length) {
        console.log(`✅ Todas as perguntas respondidas por ${numeroContato}`);
        return;
    }

    const pergunta = perguntas[indicePerguntaAtual];
    let mensagem = `${pergunta.Pergunta}\n\n`;

    for (let i = 1; i <= 5; i++) {
        if (pergunta[`Resposta${i}`]) {
            mensagem += `${i} - ${pergunta[`Resposta${i}`]}\n`;
        }
    }

    mensagem += '\nResponda com 1, 2, 3, 4 ou 5';

    const numeroComSufixo = `${numeroContato}@s.whatsapp.net`;

    try {
        await socket.sendMessage(numeroComSufixo, { text: mensagem });
        console.log(`✅ Pergunta enviada para ${numeroContato}`);
    } catch (erro) {
        console.error(`❌ Falha ao enviar pergunta para ${numeroContato}: ${erro.message}`);
    }
}

// Função para tratar uma resposta recebida
async function tratarResposta(socket, mensagem) {
    const numeroContato = mensagem.key.remoteJid.replace('@s.whatsapp.net', '');
    const contato = contatos.find(c => c.Numero === numeroContato);

    if (!contato) {
        console.log(`❌ Contato desconhecido: ${numeroContato}`);
        return;
    }

    // Se a mensagem for mídia, ignorar e informar o usuário
    if (mensagem.message.audioMessage || mensagem.message.imageMessage || mensagem.message.videoMessage || mensagem.message.stickerMessage) {
        console.log(`❌ Mensagem de mídia recebida de ${contato.Nome}. Ignorando...`);
        await socket.sendMessage(mensagem.key.remoteJid, { text: 'Por favor, envie apenas uma resposta numérica entre 1 e 5.' });
        return;
    }

    const resposta = mensagem.message.conversation?.trim();
    if (!resposta) return;

    const indicePerguntaAtual = progresso[numeroContato];
    if (indicePerguntaAtual >= perguntas.length) {
        console.log(`❌ ${contato.Nome} já respondeu todas as perguntas.`);
        await socket.sendMessage(mensagem.key.remoteJid, { text: 'Você já respondeu todas as perguntas. Obrigado!' });
        return;
    }

    const pergunta = perguntas[indicePerguntaAtual];
    const indiceResposta = parseInt(resposta, 10);

    if (!isNaN(indiceResposta) && indiceResposta >= 1 && indiceResposta <= 5) {
        const respostaSelecionada = pergunta[`Resposta${indiceResposta}`];

        if (respostaSelecionada) {
            console.log(`✅ Resposta válida recebida de ${contato.Nome}: ${respostaSelecionada}`);

            const linhaPergunta = respostas.findIndex(linha => linha[0] === pergunta.Pergunta);
            const colunaContato = respostas[0].indexOf(contato.Nome);

            if (linhaPergunta === -1) {
                const novaLinha = Array(respostas[0].length).fill('');
                novaLinha[0] = pergunta.Pergunta;
                novaLinha[colunaContato] = respostaSelecionada;
                respostas.push(novaLinha);
            } else {
                respostas[linhaPergunta][colunaContato] = respostaSelecionada;
            }

            const novaPlanilhaRespostas = xlsx.utils.aoa_to_sheet(respostas);
            respostasWorkbook.Sheets['Respostas'] = novaPlanilhaRespostas;
            xlsx.writeFile(respostasWorkbook, './pesquisa.xlsx');

            await socket.sendMessage(mensagem.key.remoteJid, { text: `Obrigado por sua resposta: ${respostaSelecionada}` });

            progresso[numeroContato]++;
            enviarPergunta(socket, numeroContato);
        } else {
            console.log(`❌ Índice de resposta inválido: ${indiceResposta}`);
        }
    } else {
        console.log(`❌ Resposta inválida de ${contato.Nome}: ${resposta}`);
        await socket.sendMessage(mensagem.key.remoteJid, { text: 'Por favor, responda com um número entre 1 e 5.' });
    }
}

// Iniciar o bot
conectarWhatsApp().catch(erro => {
    console.error('Erro ao conectar ao WhatsApp:', erro);
});
