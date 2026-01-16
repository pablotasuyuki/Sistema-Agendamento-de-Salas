import {
  Client, GatewayIntentBits, Partials, ActionRowBuilder, ButtonBuilder, ButtonStyle,
  Events, ModalBuilder, TextInputBuilder, TextInputStyle, StringSelectMenuBuilder, EmbedBuilder, SlashCommandBuilder, REST, Routes
} from 'discord.js';
import dotenv from 'dotenv';
import fs from 'fs';
import { v4 as uuidv4 } from 'uuid';
import ExcelJS from 'exceljs';
import cron from 'node-cron';
import nodemailer from 'nodemailer';

dotenv.config();

const TOKEN = process.env.TOKEN?.trim();
const CLIENT_ID = process.env.CLIENT_ID?.trim();
const FILE_DB = 'agendamentos_salas.json';

if (!TOKEN || !CLIENT_ID) {
  console.error('❌ Variáveis de ambiente ausentes.');
  process.exit(1);
}

// AGENDAMENTOS
const agendamentosSalas = [];
const agendamentosEmProgresso = new Map();

function salvarAgendamentosSalas() {
  try {
    fs.writeFileSync(FILE_DB, JSON.stringify(agendamentosSalas, null, 2));
  } catch (err) {
    console.error('Erro ao salvar agendamentos:', err);
  }
}
function carregarAgendamentosSalas() {
  if (fs.existsSync(FILE_DB)) {
    try {
      const raw = fs.readFileSync(FILE_DB, 'utf8');
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) agendamentosSalas.push(...parsed);
    } catch (err) {
      console.error('Erro ao carregar agendamentos:', err);
    }
  }
}
carregarAgendamentosSalas();

function normalizarDataAgendamento(dataInput) {
  const cleaned = dataInput.replace(/\s/g, '');
  let match = cleaned.match(/^(\d{2})[\/]?(\d{2})[\/]?(\d{4})$/);
  if (!match) return null;
  const dia = parseInt(match[1], 10);
  const mes = parseInt(match[2], 10);
  const ano = parseInt(match[3], 10);
  if (dia < 1 || dia > 31 || mes < 1 || mes > 12 || ano < 2024 || ano > 2099) return null;
  const data = new Date(ano, mes - 1, dia);
  const agora = new Date();
  agora.setHours(0,0,0,0);
  if (isNaN(data.getTime()) || data < agora) return null;
  return `${String(dia).padStart(2,'0')}/${String(mes).padStart(2,'0')}/${ano}`;
}
function normalizarHorarioAgendamento(horarioInput) {
  const cleaned = horarioInput.replace(/\s/g, '');
  let [inicio, fim] = cleaned.split('-');
  if (!inicio || !fim) return null;
  inicio = inicio.includes(':') ? inicio : `${inicio.padStart(2, '0')}:00`;
  fim = fim.includes(':') ? fim : `${fim.padStart(2, '0')}:00`;
  if (!/^\d{2}:\d{2}$/.test(inicio) || !/^\d{2}:\d{2}$/.test(fim)) return null;
  const iniM = parseInt(inicio.slice(0,2))*60 + parseInt(inicio.slice(3,5));
  const fimM = parseInt(fim.slice(0,2))*60 + parseInt(fim.slice(3,5));
  if (fimM <= iniM) return null;
  return `${inicio}-${fim}`;
}
function toMin(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + (m || 0);
}
function usuarioPodeExportar(interaction) {
  try {
    const roleId = '1294332790086566032';
    const roles = interaction.member?.roles;
    if (!roles) return false;
    if (Array.isArray(roles)) return roles.includes(roleId);
    if (roles.cache) return roles.cache.has(roleId);
    if (typeof roles === 'object') return Object.values(roles).includes(roleId);
  } catch (err) {}
  return false;
}
function horarioConflitante(sala, data, horarioNovo) {
  const [inicioNovo, fimNovo] = horarioNovo.includes('-') ? horarioNovo.split('-') : [horarioNovo, horarioNovo];
  const minInicioNovo = toMin(inicioNovo);
  const minFimNovo = toMin(fimNovo);
  return agendamentosSalas.some(a => {
    if (a.sala !== sala || a.data !== data) return false;
    if (a.status === 'Cancelada') return false;
    const [inicio, fim] = a.horario.includes('-') ? a.horario.split('-') : [a.horario, a.horario];
    const minInicio = toMin(inicio);
    const minFim = toMin(fim);
    return minInicioNovo < minFim && minFimNovo > minInicio;
  });
}
function parseDateTime(data, horario) {
  if (!data || !horario) return null;
  const [dia, mes, ano] = data.split('/');
  const [horaInicio] = horario.split('-');
  const [h, m] = horaInicio.split(':');
  return new Date(parseInt(ano), parseInt(mes)-1, parseInt(dia), parseInt(h), parseInt(m));
}
function getMesAnoOptions() {
  const meses = new Set();
  const now = new Date();
  meses.add(`${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`);
  for (const ag of agendamentosSalas) {
    if (ag.data) {
      const parts = ag.data.split('/');
      if (parts.length === 3) {
        const mes = parts[1];
        const ano = parts[2];
        meses.add(`${mes}/${ano}`);
      }
    }
  }
  const arr = Array.from(meses).sort((a, b) => {
    const [ma, aa] = a.split('/').map(Number);
    const [mb, ab] = b.split('/').map(Number);
    if (aa !== ab) return aa - ab;
    return ma - mb;
  });
  return arr.map(mesano => {
    const [mes, ano] = mesano.split('/');
    return { label: `${mes}/${ano}`, value: mesano };
  });
}
function filtrarAgendamentosPorMesAno(agendamentos, mesano, userId=null, onlyMine=false) {
  const [mes, ano] = mesano.split('/');
  return agendamentos.filter(a => {
    if (!a.data) return false;
    if (a.status === 'Cancelada') return false;
    const parts = a.data.split('/');
    if (parts[1] !== mes || parts[2] !== ano) return false;
    if (onlyMine && userId) {
      const responsavelId = a.responsavelId || a.responsavel;
      const usuarioId = a.usuarioId || a.usuario;
      if (responsavelId === userId || usuarioId === userId) return true;
      if (Array.isArray(a.participantes) && a.participantes.some(p => p.id === userId)) return true;
      return false;
    }
    return true;
  });
}
const paginacaoCalendario = new Map();

async function checarEEnviarAvisos(client) {
  const agora = new Date();
  for (const ag of agendamentosSalas) {
    try {
      if (ag.status === 'Cancelada') continue;
      if (!ag.participantes || ag.participantes.length === 0) continue;
      if (!ag.data || !ag.horario) continue;
      const dtReuniao = parseDateTime(ag.data, ag.horario);
      if (!dtReuniao) continue;
      const diffMs = dtReuniao.getTime() - agora.getTime();
      const diffHoras = diffMs / (1000 * 60 * 60);

      if (!ag.notificado1d && diffHoras <= 24 && diffHoras > 1) {
        ag.notificado1d = true;
        salvarAgendamentosSalas();
        for (const participante of ag.participantes) {
          try {
            const user = await client.users.fetch(participante.id);
            await user.send(
              `🔔 Lembrete (24h): Você tem uma reunião agendada para **${ag.data}** na sala **${ag.sala}** às **${ag.horario}**!\n\nTítulo: ${ag.titulo}`
            );
          } catch {}
        }
      }

      if (!ag.notificado1h && diffHoras <= 1 && diffHoras > 0) {
        ag.notificado1h = true;
        salvarAgendamentosSalas();
        for (const participante of ag.participantes) {
          try {
            const user = await client.users.fetch(participante.id);
            await user.send(
              `🔔 Lembrete (1h): Falta 1 hora para sua reunião na **${ag.sala}** em **${ag.data}** às **${ag.horario}**!\n\nTítulo: ${ag.titulo}`
            );
          } catch {}
        }
      }
    } catch (err) {
      console.error('Erro ao checar avisos para agendamento', ag.id, err);
    }
  }
}
const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildMembers
  ],
  partials: [Partials.Channel],
});
client.once(Events.ClientReady, () => {
  console.log(`🤖 Bot online como ${client.user.tag}`);
});
setInterval(() => {
  if (client.isReady && client.isReady()) checarEEnviarAvisos(client);
}, 60*1000);

const commands = [
  new SlashCommandBuilder().setName('menu-salas').setDescription('Abrir menu de agendamento de salas de reunião')
].map(cmd => cmd.toJSON());
const rest = new REST({ version: '10' }).setToken(TOKEN);
(async () => {
  try {
    await rest.put(Routes.applicationCommands(CLIENT_ID), { body: commands });
    console.log('✅ Comandos de barra registrados!');
  } catch (error) {
    console.error('❌ Falha ao registrar comandos:', error);
    process.exit(1);
  }
})();

async function replyMenuSalas(interactionOrMessage) {
  const embed = new EmbedBuilder()
    .setColor(0x2C3E50)
    .setTitle('Agendamento de Salas de Reunião')
    .setImage("https://media.discordapp.net/attachments/1372944028458418217/1428418532957491291/Sala_de_reuniao.jpg?ex=68f26dec&is=68f11c6c&hm=90c8a5e6544b28ed74eb78f5d6e224993c90d993a9be49f2a43dc47b28ccbe6e&=&format=webp&width=1163&height=960");

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('agendar_sala').setLabel('Agendar Sala').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('meus_agendamentos_sala').setLabel('Meus agendamentos').setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId('calendario_agendamentos_sala').setLabel('Calendário').setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId('cancelar_agendamento_sala').setLabel('Cancelar agendamento').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('exportar_agendamentos_sala').setLabel('Exportar Dados').setStyle(ButtonStyle.Success)
  );
  if ('replied' in interactionOrMessage || 'deferred' in interactionOrMessage) {
    if (interactionOrMessage.replied || interactionOrMessage.deferred) {
      await interactionOrMessage.followUp({ embeds: [embed], components: [row] }); // sem flags: 64
    } else {
      await interactionOrMessage.reply({ embeds: [embed], components: [row] }); // sem flags: 64
    }
  } else {
    await interactionOrMessage.reply({ embeds: [embed], components: [row] }); // sem flags: 64
  }
}
client.on(Events.MessageCreate, async message => {
  if (message.content === '!menu-salas') await replyMenuSalas(message);
});

client.on(Events.InteractionCreate, async interaction => {
  try {
    // MENU PRINCIPAL
    if (interaction.isChatInputCommand()) {
      if (interaction.commandName === 'menu-salas') return await replyMenuSalas(interaction);
    }

    // MEUS AGENDAMENTOS
    if (interaction.isButton() && interaction.customId === 'meus_agendamentos_sala') {
      const options = getMesAnoOptions();
      const selectRow = new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId('select_mes_meus_agendamentos')
          .setPlaceholder('Selecione o mês dos seus agendamentos')
          .addOptions(options)
      );
      await interaction.reply({
        content: 'Selecione o mês para exibir seus agendamentos:',
        components: [selectRow],
        flags: 64
      });
      return;
    }
    // CALENDÁRIO
    if (interaction.isButton() && interaction.customId === 'calendario_agendamentos_sala') {
      const options = getMesAnoOptions();
      const selectRow = new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId('select_mes_calendario')
          .setPlaceholder('Selecione o mês para ver o calendário')
          .addOptions(options)
      );
      await interaction.reply({
        content: 'Selecione o mês do calendário de agendamentos:',
        components: [selectRow],
        flags: 64
      });
      return;
    }
    // CANCELAMENTO (SEGURO!)
    if (interaction.isButton() && interaction.customId === 'cancelar_agendamento_sala') {
      const meus = agendamentosSalas.filter(a =>
        a.status !== 'Cancelada' &&
        (
          a.responsavelId === interaction.user.id ||
          a.usuarioId === interaction.user.id ||
          (Array.isArray(a.participantes) && a.participantes.some(p => p.id === interaction.user.id))
        )
      );
      if (meus.length === 0) {
        await interaction.reply({
          content: 'Você não possui agendamentos para cancelar.',
          flags: 64
        });
        return;
      }
      const mesesAnosSet = new Set(
        meus.map(a => {
          const [dia, mes, ano] = a.data.split('/');
          return `${mes}/${ano}`;
        })
      );
      const mesesAnos = Array.from(mesesAnosSet).sort((a, b) => {
        const [mA, yA] = a.split('/').map(Number);
        const [mB, yB] = b.split('/').map(Number);
        return yA !== yB ? yA - yB : mA - mB;
      });
      const select = new StringSelectMenuBuilder()
        .setCustomId('cancelar_sala_select_mes_ano')
        .setPlaceholder('Selecione o mês e ano')
        .addOptions(mesesAnos.map(ma => {
          const [mes, ano] = ma.split('/');
          return { label: `${mes.padStart(2, '0')}/${ano}`, value: ma }
        }));
      const row = new ActionRowBuilder().addComponents(select);
      await interaction.reply({
        content: 'Selecione o mês e ano dos agendamentos que deseja cancelar:',
        components: [row],
        flags: 64
      });
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId === 'cancelar_sala_select_mes_ano') {
      const [mes, ano] = interaction.values[0].split('/');
      const meus = agendamentosSalas.filter(a =>
        a.status !== 'Cancelada' &&
        a.data?.split('/')[1] === mes &&
        a.data?.split('/')[2] === ano &&
        (
          a.responsavelId === interaction.user.id ||
          a.usuarioId === interaction.user.id ||
          (Array.isArray(a.participantes) && a.participantes.some(p => p.id === interaction.user.id))
        )
      );
      if (meus.length === 0) {
        try {
          await interaction.update({
            content: 'Nenhum agendamento encontrado neste mês/ano para cancelar.',
            components: [],
          });
        } catch (err) {}
        return;
      }
      const rows = [];
      for (let i = 0; i < meus.length; i += 5) {
        rows.push(
          new ActionRowBuilder().addComponents(
            ...meus.slice(i, i + 5).map(ag =>
              new ButtonBuilder()
                .setCustomId(`cancelar_sala_${ag.id}`)
                .setLabel(`${ag.data} ${ag.horario} - ${ag.sala}`)
                .setStyle(ButtonStyle.Danger)
            )
          )
        );
      }
      try {
        await interaction.update({
          content: 'Selecione qual agendamento deseja cancelar:',
          components: rows,
        });
      } catch (err) {}
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith('cancelar_sala_')) {
      const id = interaction.customId.replace('cancelar_sala_', '');
      const idx = agendamentosSalas.findIndex(a =>
        a.id === id &&
        a.status !== 'Cancelada' &&
        (
          a.responsavelId === interaction.user.id ||
          a.usuarioId === interaction.user.id ||
          (Array.isArray(a.participantes) && a.participantes.some(p => p.id === interaction.user.id))
        )
      );
      if (idx !== -1) {
        agendamentosSalas[idx].status = 'Cancelada';
        salvarAgendamentosSalas();
        try {
          await interaction.update({
            content: '✅ Agendamento cancelado com sucesso.',
            components: [],
          });
        } catch (err) {}
        return;
      }
      try {
        await interaction.update({
          content: 'Agendamento não encontrado ou já cancelado.',
          components: [],
        });
      } catch (err) {}
      return;
    }

    // EXPORTAÇÃO
    if (interaction.isButton() && interaction.customId === 'exportar_agendamentos_sala') {
      if (!usuarioPodeExportar(interaction)) {
        await interaction.reply({
          content: 'Você não tem permissão para exportar agendamentos.',
          flags: 64
        });
        return;
      }
      if (agendamentosSalas.length === 0) {
        await interaction.reply({
          content: 'Não há agendamentos para exportar.',
          flags: 64
        });
        return;
      }
      const workbook = new ExcelJS.Workbook();
      const ws = workbook.addWorksheet('Agendamentos');
      ws.columns = [
        { header: 'Data', key: 'data', width: 12 },
        { header: 'Horário', key: 'horario', width: 16 },
        { header: 'Sala', key: 'sala', width: 16 },
        { header: 'Título', key: 'titulo', width: 40 },
        { header: 'Responsável (nome)', key: 'responsavel', width: 20 },
        { header: 'Responsável (id)', key: 'responsavelId', width: 20 },
        { header: 'Usuário (nome)', key: 'usuario', width: 20 },
        { header: 'Usuário (id)', key: 'usuarioId', width: 20 },
        { header: 'Participantes (tags)', key: 'participantes', width: 40 },
        { header: 'Status', key: 'status', width: 14 }
      ];
      for (const ag of agendamentosSalas) {
        ws.addRow({
          data: ag.data,
          horario: ag.horario,
          sala: ag.sala,
          titulo: ag.titulo,
          responsavel: ag.responsavel || '',
          responsavelId: ag.responsavelId || '',
          usuario: ag.usuario || '',
          usuarioId: ag.usuarioId || '',
          participantes: (ag.participantes?.map(p => `<@${p.id}>`).join(', ')) || '',
          status: ag.status || ''
        });
      }
      const buffer = await workbook.xlsx.writeBuffer();
      await interaction.reply({
        content: `Exportação de todo o histórico de agendamentos:`,
        files: [{ attachment: Buffer.from(buffer), name: `agendamentos_completo.xlsx` }],
        embeds: [],
        components: [],
        flags: 64
      });
      return;
    }
    // MEUS AGENDAMENTOS
    if (interaction.isStringSelectMenu() && interaction.customId === 'select_mes_meus_agendamentos') {
      const mesano = interaction.values[0];
      const meus = filtrarAgendamentosPorMesAno(agendamentosSalas, mesano, interaction.user.id, true);
      if (meus.length === 0) {
        await interaction.update({ content: 'Você não possui agendamentos neste mês.', embeds: [], components: [] });
        return;
      }
      const embed = new EmbedBuilder()
        .setTitle(`Seus agendamentos (${mesano})`)
        .setColor(0x3498db)
        .setDescription(
          meus.map(ag =>
            `**${ag.data} ${ag.horario}** - ${ag.sala}\nTítulo: ${ag.titulo}\nStatus: ${ag.status}\n---`
          ).join('\n')
        );
      await interaction.update({ embeds: [embed], content: '', components: [] });
      return;
    }
    // CALENDÁRIO
    if (interaction.isStringSelectMenu() && interaction.customId === 'select_mes_calendario') {
      const mesano = interaction.values[0];
      const filtrados = filtrarAgendamentosPorMesAno(agendamentosSalas, mesano);
      if (filtrados.length === 0) {
        await interaction.update({ content: 'Não há agendamentos neste mês.', embeds: [], components: [] });
        return;
      }
      const embeds = filtrados
        .sort((a, b) => {
          const dtA = parseDateTime(a.data, a.horario);
          const dtB = parseDateTime(b.data, b.horario);
          return dtA - dtB;
        })
        .map(ag =>
          new EmbedBuilder()
            .setTitle(`${ag.data} ${ag.horario} - ${ag.sala}`)
            .setColor(0x27ae60)
            .setDescription(
              `**Título:** ${ag.titulo}\n` +
              `**Responsável:** ${ag.responsavel || 'Desconhecido'}\n` +
              `**Status:** ${ag.status || '—'}\n` +
              `**Participantes:** ${ag.participantes?.map(p => `<@${p.id}>`).join(', ') || 'Nenhum'}`
            )
        );
      const totalPages = Math.ceil(embeds.length / 10);
      const page = totalPages - 1;
      paginacaoCalendario.set(interaction.user.id, { embeds, page, mesano });

      const start = page * 10;
      const end = start + 10;
      const pageEmbeds = embeds.slice(start, end);

      const row = new ActionRowBuilder();
      if (page > 0) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_anteriores_salas')
            .setLabel('Ver anteriores')
            .setStyle(ButtonStyle.Primary)
        );
      }
      if (page < totalPages - 1) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_proximos_salas')
            .setLabel('Ver próximos')
            .setStyle(ButtonStyle.Secondary)
        );
      }

      await interaction.update({
        embeds: pageEmbeds,
        content: '',
        components: row.components.length > 0 ? [row] : [],
      });
      return;
    }
    if (interaction.isButton() && interaction.customId === 'ver_anteriores_salas') {
      const paginacao = paginacaoCalendario.get(interaction.user.id);
      if (!paginacao) {
        await interaction.reply({ content: "Navegação expirada ou não encontrada.", flags: 64 });
        return;
      }
      let page = paginacao.page - 1;
      if (page < 0) page = 0;
      paginacao.page = page;
      const start = page * 10;
      const end = start + 10;
      const pageEmbeds = paginacao.embeds.slice(start, end);
      const row = new ActionRowBuilder();
      if (page > 0) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_anteriores_salas')
            .setLabel('Ver anteriores')
            .setStyle(ButtonStyle.Primary)
        );
      }
      if (page < Math.ceil(paginacao.embeds.length / 10) - 1) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_proximos_salas')
            .setLabel('Ver próximos')
            .setStyle(ButtonStyle.Secondary)
        );
      }
      await interaction.update({
        embeds: pageEmbeds,
        content: '',
        components: row.components.length > 0 ? [row] : [],
      });
      return;
    }
    if (interaction.isButton() && (interaction.customId === 'ver_anteriores_salas' || interaction.customId === 'ver_proximos_salas')) {
      const paginacao = paginacaoCalendario.get(interaction.user.id);
      if (!paginacao) {
        await interaction.reply({ content: "Navegação expirada ou não encontrada.", flags: 64 });
        return;
      }
      const totalPages = Math.ceil(paginacao.embeds.length / 10);
      let page = paginacao.page;
      if (interaction.customId === 'ver_anteriores_salas') page--;
      if (interaction.customId === 'ver_proximos_salas') page++;
      page = Math.max(0, Math.min(totalPages - 1, page));
      paginacao.page = page;
      const start = page * 10;
      const end = start + 10;
      const pageEmbeds = paginacao.embeds.slice(start, end);
      const row = new ActionRowBuilder();
      if (page > 0) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_anteriores_salas')
            .setLabel('Ver anteriores')
            .setStyle(ButtonStyle.Primary)
        );
      }
      if (page < totalPages - 1) {
        row.addComponents(
          new ButtonBuilder()
            .setCustomId('ver_proximos_salas')
            .setLabel('Ver próximos')
            .setStyle(ButtonStyle.Secondary)
        );
      }
      await interaction.update({
        embeds: pageEmbeds,
        content: '',
        components: row.components.length > 0 ? [row] : [],
      });
      return;
    }
    // AGENDAMENTO
    if (interaction.isButton() && interaction.customId === 'agendar_sala') {
      const modal = new ModalBuilder()
        .setCustomId('modal_data_horario_agendamento_sala')
        .setTitle('Agendamento de Sala - Data e Horário')
        .addComponents(
          new ActionRowBuilder().addComponents(
            new TextInputBuilder()
              .setCustomId('data')
              .setLabel('Para qual data você precisa agendar?')
              .setStyle(TextInputStyle.Short)
              .setPlaceholder('DD/MM/AAAA ou DDMMAAAA')
              .setRequired(true)
          ),
          new ActionRowBuilder().addComponents(
            new TextInputBuilder()
              .setCustomId('horario')
              .setLabel('Horário inicial e final (ex: 09:00-11:00)')
              .setStyle(TextInputStyle.Short)
              .setRequired(true)
              .setPlaceholder('09:00-11:00')
          )
        );
      await interaction.showModal(modal);
      return;
    }
    if (interaction.isModalSubmit() && interaction.customId === 'modal_data_horario_agendamento_sala') {
      const rawData = interaction.fields.getTextInputValue('data').trim();
      const dataSelecionada = normalizarDataAgendamento(rawData);
      if (!dataSelecionada) {
        await interaction.reply({
          content: 'Data inválida! Use datas futuras no formato DD/MM/AAAA ou DDMMAAAA (dia ≤ 31, mês ≤ 12, ano entre 2024 e 2099).',
          flags: 64
        });
        return;
      }
      const rawHorario = interaction.fields.getTextInputValue('horario').trim();
      const horario = normalizarHorarioAgendamento(rawHorario);
      if (!horario) {
        await interaction.reply({
          content: 'Horário inválido! Informe um horário inicial e um final. Exemplos válidos: 09:00-11:00, 08-17.',
          flags: 64
        });
        return;
      }
      const todasSalas = [
        { nome: 'Sala Grande', id: 'select_sala_Sala Grande' },
        { nome: 'Sala Menor', id: 'select_sala_Sala Menor' },
        { nome: 'Sala Menor C/Mesa', id: 'select_sala_Sala Menor C/Mesa' }
      ];
      const salasDisponiveis = todasSalas.filter(sala =>
        !horarioConflitante(sala.nome, dataSelecionada, horario)
      );
      if (salasDisponiveis.length === 0) {
        await interaction.reply({
          content: `Nenhuma sala está disponível para ${dataSelecionada} no horário ${horario}. Tente outra data ou horário.`,
          flags: 64
        });
        return;
      }
      agendamentosEmProgresso.set(interaction.user.id, { data: dataSelecionada, horario });
      const row = new ActionRowBuilder().addComponents(
        ...salasDisponiveis.map(sala =>
          new ButtonBuilder().setCustomId(sala.id).setLabel(sala.nome).setStyle(ButtonStyle.Primary)
        )
      );
      await interaction.reply({
        content: `Escolha a sala disponível para **${dataSelecionada}** no horário **${horario}**:`,
        components: [row],
        flags: 64
      });
      return;
    }
    if (interaction.isButton() && interaction.customId.startsWith('select_sala_')) {
      const sala = interaction.customId.replace('select_sala_', '');
      const progresso = agendamentosEmProgresso.get(interaction.user.id);
      if (!progresso) {
        await interaction.reply({ content: 'Erro: passo anterior não encontrado. Tente novamente.', flags: 64 });
        return;
      }
      progresso.sala = sala;
      const modal = new ModalBuilder()
        .setCustomId('modal_titulo_agendamento_sala')
        .setTitle('Título/Descrição da Reunião')
        .addComponents(
          new ActionRowBuilder().addComponents(
            new TextInputBuilder()
              .setCustomId('titulo')
              .setLabel('Título ou descrição da reunião')
              .setStyle(TextInputStyle.Short)
              .setRequired(true)
              .setPlaceholder('Ex: Reunião de equipe...')
          )
        );
      await interaction.showModal(modal);
      return;
    }
    if (interaction.isModalSubmit() && interaction.customId === 'modal_titulo_agendamento_sala') {
      const progresso = agendamentosEmProgresso.get(interaction.user.id);
      if (!progresso) {
        await interaction.reply({ content: 'Erro: passo anterior não encontrado. Tente novamente.', flags: 64 });
        return;
      }
      progresso.titulo = interaction.fields.getTextInputValue('titulo').trim();
      const guild = await client.guilds.fetch(interaction.guildId);
      await guild.members.fetch();
      const allowedRoles = ['1371460014652391514', '1371460180918665378'];
      const nomesRemover = [
        "andré rocha | T.I",
        "pedro gabriel - arquivo",
        "etienne - arquivo"
      ];
      const allowedMembers = guild.members.cache.filter(member =>
        (member.roles.cache.has(allowedRoles[0]) || member.roles.cache.has(allowedRoles[1])) &&
        !nomesRemover.some(rem =>
          (member.displayName || '').trim().toLowerCase() === rem.trim().toLowerCase() ||
          (member.user.username || '').trim().toLowerCase() === rem.trim().toLowerCase()
        )
      );
      const allowedMembersArr = Array.from(allowedMembers.values())
        .sort((a, b) => a.displayName.localeCompare(b.displayName, 'pt-BR', { sensitivity: 'base' }));
      if (allowedMembersArr.length === 0) {
        await interaction.reply({ content: 'Nenhum membro elegível encontrado para reunião.', flags: 64 });
        return;
      }
      const selectRows = [];
      const maxPerMenu = 25;
      const totalMenus = Math.ceil(allowedMembersArr.length / maxPerMenu);
      for (let i = 0; i < totalMenus; i++) {
        const membersSlice = allowedMembersArr.slice(i * maxPerMenu, (i + 1) * maxPerMenu);
        const selectMenu = new StringSelectMenuBuilder()
          .setCustomId(`select_participantes_${i}`)
          .setPlaceholder(`Participantes (${i * maxPerMenu + 1}-${i * maxPerMenu + membersSlice.length})`)
          .setMinValues(0)
          .setMaxValues(membersSlice.length)
          .addOptions(
            membersSlice.map(member => ({
              label: member.displayName,
              value: member.id,
              description: member.user.username
            }))
          );
        selectRows.push(new ActionRowBuilder().addComponents(selectMenu));
      }
      const actionRow = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId('concluir_participantes')
          .setLabel('Concluir seleção')
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId('sem_participantes')
          .setLabel('Não adicionar participantes')
          .setStyle(ButtonStyle.Secondary)
      );
      progresso.participantesParciais = [];
      progresso.totalMenus = totalMenus;
      await interaction.reply({
        content: `Selecione os participantes em todos os menus que quiser.\nQuando terminar, clique em "Concluir seleção".\nSe não quiser adicionar ninguém, clique em "Não adicionar participantes".`,
        components: [...selectRows, actionRow],
        flags: 64
      });
      return;
    }
    if (interaction.isStringSelectMenu() && interaction.customId.startsWith('select_participantes_')) {
      const progresso = agendamentosEmProgresso.get(interaction.user.id);
      if (!progresso) {
        await interaction.reply({ content: 'Erro: passo anterior não encontrado. Tente novamente.', flags: 64 });
        return;
      }
      const menuIndex = parseInt(interaction.customId.replace('select_participantes_', ''));
      progresso.participantesParciais = progresso.participantesParciais || [];
      progresso.participantesParciais[menuIndex] = interaction.values;
      await interaction.reply({ content: 'Participantes selecionados neste menu! Se quiser, selecione em outros menus ou conclua.', flags: 64 });
      return;
    }
    if (interaction.isButton() && (interaction.customId === 'concluir_participantes' || interaction.customId === 'sem_participantes')) {
      const progresso = agendamentosEmProgresso.get(interaction.user.id);
      if (!progresso) {
        await interaction.reply({ content: 'Erro: passo anterior não encontrado. Tente novamente.', flags: 64 });
        return;
      }
      let selecionados = [];
      if (interaction.customId === 'concluir_participantes') {
        if (Array.isArray(progresso.participantesParciais)) {
          progresso.participantesParciais.forEach(arr => {
            if (Array.isArray(arr)) arr.forEach(id => { if (!selecionados.includes(id)) selecionados.push(id); });
          });
        }
        progresso.participantes = selecionados.map(userId => {
          const member = interaction.guild.members.cache.get(userId);
          if (member)
            return { id: member.id, tag: member.user.tag, username: member.user.username };
          return null;
        }).filter(Boolean);
      } else {
        progresso.participantes = [];
      }
      if (horarioConflitante(progresso.sala, progresso.data, progresso.horario)) {
        await interaction.reply({ content: 'Já existe uma reserva para esse horário nesta sala!', flags: 64 });
        return;
      }
      progresso.status = 'Agendada';
      const avatarURL = interaction.user.displayAvatarURL();
      const novoAgendamento = {
        id: uuidv4(),
        data: progresso.data,
        sala: progresso.sala,
        horario: progresso.horario,
        responsavel: interaction.member?.nickname || interaction.member?.displayName || interaction.user.globalName || interaction.user.username,
        responsavelId: interaction.user.id,
        usuario: interaction.user.username,
        usuarioId: interaction.user.id,
        participantes: progresso.participantes || [],
        status: 'Agendada',
        avatar: avatarURL,
        presencas: {},
        titulo: progresso.titulo
      };
      agendamentosSalas.push(novoAgendamento);
      salvarAgendamentosSalas();
      if (novoAgendamento.participantes.length > 0) {
        novoAgendamento.participantes.forEach(async userObj => {
          try {
            const user = await client.users.fetch(userObj.id);
            const embed = new EmbedBuilder()
              .setTitle('Agendamento da Reunião')
              .setColor(0x2ecc71)
              .setDescription(
                `**Data:** ${progresso.data}\n` +
                `**Sala:** ${progresso.sala}\n` +
                `**Horário:** ${progresso.horario}\n` +
                `**Título:** ${progresso.titulo}\n` +
                `**Responsável:** <@${novoAgendamento.responsavelId}>\n` +
                `**Participantes:** ${novoAgendamento.participantes.map(p => `<@${p.id}>`).join(', ') || 'Nenhum'}`
              );
            const row = new ActionRowBuilder().addComponents(
  new ButtonBuilder()
    .setCustomId(`confirmar_presenca_${novoAgendamento.id}`)
    .setLabel('Confirmar presença')
    .setStyle(ButtonStyle.Success),
  new ButtonBuilder()
    .setCustomId(`reprovar_presenca_${novoAgendamento.id}`)
    .setLabel('Reprovar presença')
    .setStyle(ButtonStyle.Danger),
  new ButtonBuilder()
    .setCustomId(`status_presenca_${novoAgendamento.id}`)
    .setLabel('Status de Presença')
    .setStyle(ButtonStyle.Secondary)
);
await user.send({
  embeds: [embed],
  components: [row]
});
          } catch {}
        });
      }
      const embed = new EmbedBuilder()
        .setTitle('Sala agendada!')
        .setAuthor({ name: novoAgendamento.responsavel, iconURL: avatarURL })
        .setDescription(
          `**Data:** ${progresso.data}\n` +
          `**Sala:** ${progresso.sala}\n` +
          `**Horário:** ${progresso.horario}\n` +
          `**Título:** ${progresso.titulo}\n` +
          `**Participantes:** ${novoAgendamento.participantes.map(p => `<@${p.id}>`).join(', ') || 'Nenhum'}\n` +
          `**Status:** Agendada`
        )
        .setColor(0x3498db)
        .setFooter({ text: 'Os participantes receberão convite para confirmar presença.' });
      await interaction.reply({ embeds: [embed], components: [row] });
      agendamentosEmProgresso.delete(interaction.user.id);
      return;
    }
    // PRESENÇA
    if (
  interaction.isButton() &&
  interaction.customId.startsWith('status_presenca_')
) {
  const id = interaction.customId.split('_').pop();
  const agendamento = agendamentosSalas.find(a => a.id === id);
  if (!agendamento) {
    await interaction.reply({ content: 'Agendamento não encontrado.', ephemeral: true });
    return;
  }
  const confirmados = agendamento.participantes.filter(p => agendamento.presencas?.[p.id] === 'Confirmada');
  const reprovados = agendamento.participantes.filter(p => agendamento.presencas?.[p.id] === 'Reprovada');
  const embed = new EmbedBuilder()
    .setTitle(`Presenças para reunião ${agendamento.sala} (${agendamento.data} - ${agendamento.horario})`)
    .setDescription(
      `**Confirmados:**\n${confirmados.map(p => `<@${p.id}>`).join('\n') || 'Ninguém ainda'}\n\n**Reprovados:**\n${reprovados.map(p => `<@${p.id}>`).join('\n') || 'Ninguém'}`
    )
    .setColor(0x27ae60);
  await interaction.reply({ embeds: [embed], ephemeral: true }); // components removido!
  return;
}

// PRESENÇA - CONFIRMAR/REPROVAR
if (
  interaction.isButton() &&
  (interaction.customId.startsWith('confirmar_presenca_') || interaction.customId.startsWith('reprovar_presenca_'))
) {
  const isConfirm = interaction.customId.startsWith('confirmar_presenca_');
  const id = interaction.customId.split('_').pop();
  const agendamento = agendamentosSalas.find(a => a.id === id);
  if (!agendamento) {
    await interaction.reply({ content: 'Agendamento não encontrado.', flags: 64 });
    return;
  }
  if (!agendamento.presencas) agendamento.presencas = {};
  if (agendamento.presencas[interaction.user.id]) {
    await interaction.reply({
      content: 'Você já registrou sua presença para esta reunião.',
      flags: 64,
    });
    return;
  }
  agendamento.presencas[interaction.user.id] = isConfirm ? 'Confirmada' : 'Reprovada';
  salvarAgendamentosSalas();
  // Sempre inclui o botão "Status de Presença"
  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`confirmar_presenca_${agendamento.id}`)
      .setLabel('Confirmar presença')
      .setStyle(ButtonStyle.Success)
      .setDisabled(isConfirm), // desabilita se já confirmou
    new ButtonBuilder()
      .setCustomId(`reprovar_presenca_${agendamento.id}`)
      .setLabel('Reprovar presença')
      .setStyle(ButtonStyle.Danger)
      .setDisabled(!isConfirm), // desabilita caso já reprovou
    new ButtonBuilder()
      .setCustomId(`status_presenca_${agendamento.id}`)
      .setLabel('Status de Presença')
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(false)
  );
  try {
    if (interaction.message && interaction.message.editable) {
      await interaction.update({
        components: [row]
      });
    } else {
      await interaction.reply({ content: isConfirm ? 'Presença confirmada.' : 'Presença reprovada.', flags: 64 });
    }
  } catch {
    await interaction.reply({ content: isConfirm ? 'Presença confirmada.' : 'Presença reprovada.', flags: 64 });
  }
  return;
}
  } catch (err) {
    console.error('Erro no InteractionCreate:', err);
    try {
      if (interaction && !interaction.replied && !interaction.deferred) await interaction.reply({ content: 'Ocorreu um erro ao processar sua ação.', flags: 64 });
    } catch {}
  }
});

async function gerarBufferExcelCompleto() {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Agendamentos');
  ws.columns = [
    { header: 'Data', key: 'data', width: 12 },
    { header: 'Horário', key: 'horario', width: 16 },
    { header: 'Sala', key: 'sala', width: 16 },
    { header: 'Título', key: 'titulo', width: 40 },
    { header: 'Responsável (nome)', key: 'responsavel', width: 20 },
    { header: 'Responsável (id)', key: 'responsavelId', width: 20 },
    { header: 'Usuário (nome)', key: 'usuario', width: 20 },
    { header: 'Usuário (id)', key: 'usuarioId', width: 20 },
    { header: 'Participantes (tags)', key: 'participantes', width: 40 },
    { header: 'Status', key: 'status', width: 14 }
  ];
  for (const ag of agendamentosSalas) {
    ws.addRow({
      data: ag.data,
      horario: ag.horario,
      sala: ag.sala,
      titulo: ag.titulo,
      responsavel: ag.responsavel || '',
      responsavelId: ag.responsavelId || '',
      usuario: ag.usuario || '',
      usuarioId: ag.usuarioId || '',
      participantes: (ag.participantes?.map(p => `<@${p.id}>`).join(', ')) || '',
      status: ag.status || ''
    });
  }
  return await workbook.xlsx.writeBuffer();
}

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'enviaremailshagg@gmail.com',
    pass: 'zqqq ehuc nryd wlqw'
  }
});

async function enviarExcelPorEmail() {
  try {
    const buffer = await gerarBufferExcelCompleto();
    await transporter.sendMail({
      from: '"Sistema de Agendamento" <enviaremailshagg@gmail.com>',
      to: 'enviaremailshagg@gmail.com',
      subject: 'Exportação automática de agendamentos',
      text: 'Segue em anexo o arquivo Excel com todo o histórico de agendamentos.',
      attachments: [
        {
          filename: 'salas.xlsx',
          content: buffer
        }
      ]
    });
    console.log('✅ Email com Excel enviado com sucesso.');
  } catch (err) {
    console.error('❌ Erro ao enviar email automático:', err);
  }
}

// Agenda: todo dia às 8h e às 13h (ajuste timezone se necessário)
cron.schedule('0 8,13 * * *', () => {
  enviarExcelPorEmail();
}, {
  timezone: 'America/Sao_Paulo'
});

// COMANDO DE TESTE NO TERMINAL: node sala.js test
if (process.argv[2] === 'test') {
  enviarExcelPorEmail();
}

client.login(TOKEN);