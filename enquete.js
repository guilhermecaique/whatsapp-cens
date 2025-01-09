import pkg from '@whiskeysockets/baileys';
import qrcode from 'qrcode-terminal';
import fs from 'fs';
import xlsx from 'xlsx';

const { makeWASocket, makeInMemoryStore, useMultiFileAuthState, DisconnectReason } = pkg;

// Create an in-memory store
const store = makeInMemoryStore({});

// Load authentication state
const { state, saveCreds } = await useMultiFileAuthState('./auth');

// Load contacts and questions from Excel files
const contactsWorkbook = xlsx.readFile('./contatos.xlsx');
const contactsSheet = contactsWorkbook.Sheets[contactsWorkbook.SheetNames[0]];
const contacts = xlsx.utils.sheet_to_json(contactsSheet);

const questionsWorkbook = xlsx.readFile('./pesquisa.xlsx');
const questionsSheet = questionsWorkbook.Sheets['Perguntas'];
const questions = xlsx.utils.sheet_to_json(questionsSheet);

// Function to connect to WhatsApp
async function connectToWhatsApp() {
    const socket = makeWASocket({
        auth: state,
    });

    store.bind(socket.ev);

    socket.ev.on('creds.update', saveCreds);

    socket.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr) {
            console.log('⚠️ Connection lost. Please scan the new QR code:');
            qrcode.generate(qr, { small: true });
        }

        if (connection === 'open') {
            console.log('✅ Connected to WhatsApp!');
            sendQuestions(socket);
        }

        if (connection === 'close') {
            const shouldReconnect = (lastDisconnect?.error)?.output?.statusCode !== DisconnectReason.loggedOut;
            if (shouldReconnect) {
                console.log('⚠️ Connection closed. Attempting to reconnect...');
                await connectToWhatsApp();
            } else {
                console.log('❌ Logged out. Please delete the "auth" folder and restart the bot to reconnect.');
            }
        }
    });

    // Listen for incoming messages
    socket.ev.on('messages.upsert', async (msg) => {
        const message = msg.messages[0];
        if (!message.key.fromMe && message.message && message.message.conversation) {
            handleResponse(socket, message);
        }
    });
}

// Function to send list messages
async function sendQuestions(socket) {
    if (questions.length === 0) {
        console.log('No questions to send.');
        return;
    }

    for (const contact of contacts) {
        const numberWithSuffix = `${contact.Numero}@s.whatsapp.net`;

        for (const question of questions) {
            const listMessage = {
                text: 'Por favor, selecione uma opção abaixo:',
                footer: 'Pesquisa',
                title: question.Pergunta,
                buttonText: 'Escolher Opção',
                sections: [
                    {
                        title: 'Opções de Resposta',
                        rows: [
                            { title: question.Resposta1, rowId: '1' },
                            { title: question.Resposta2, rowId: '2' },
                            { title: question.Resposta3, rowId: '3' },
                            { title: question.Resposta4, rowId: '4' },
                            { title: question.Resposta5, rowId: '5' },
                        ].filter(row => row.title), // Removes empty rows
                    }
                ]
            };

            try {
                await socket.sendMessage(numberWithSuffix, { listMessage });
                console.log(`✅ List message sent to ${contact.Nome} (${contact.Numero})`);
            } catch (err) {
                console.error(`❌ Failed to send list message to ${contact.Nome} (${contact.Numero}):`, err);
            }
        }
    }
}

// Function to handle responses
async function handleResponse(socket, message) {
    const response = message.message.listResponseMessage?.singleSelectReply?.selectedRowId;
    const contactNumber = message.key.remoteJid.replace('@s.whatsapp.net', '');

    const contact = contacts.find(c => c.Numero === contactNumber);
    if (!contact) {
        console.log(`❌ Unknown contact: ${contactNumber}`);
        return;
    }

    if (response) {
        console.log(`✅ Received valid response from ${contact.Nome}: ${response}`);
        await socket.sendMessage(message.key.remoteJid, { text: `Obrigado por sua resposta: ${response}` });
    } else {
        console.log(`❌ Invalid response from ${contact.Nome}`);
        await socket.sendMessage(message.key.remoteJid, { text: 'Por favor, selecione uma opção válida.' });
    }
}

// Start the bot
connectToWhatsApp().catch(err => {
    console.error('Error connecting to WhatsApp:', err);
});
