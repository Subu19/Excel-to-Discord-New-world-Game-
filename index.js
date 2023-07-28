const Discord = require("discord.js");
const { Client, GatewayIntentBits } = require("discord.js");
require("dotenv").config();
const fs = require("fs/promises");
const config = require("./config.json");
const xlsx = require("xlsx");
const { default: axios } = require("axios");
const client = new Discord.Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildMembers,
    GatewayIntentBits.GuildModeration,
    GatewayIntentBits.GuildIntegrations,
    GatewayIntentBits.AutoModerationExecution,
    GatewayIntentBits.AutoModerationConfiguration,
    GatewayIntentBits.GuildPresences,
  ],
});

// Bot token - Replace with your own token
const token = process.env.TOKEN;

client.login(token);

client.on("ready", () => {
  console.log(`Logged in as ${client.user.tag}`);
  runtask();
});

//tasks

function runtask() {
  setInterval(() => {
    fetchPlayers();
  }, 10000);
}

async function fetchPlayers() {
  if (!config.channel) return;
  const data = await readExcel("./test.xlsx");
  const columnData = await generateProperJSON(data);
  const embed = new Discord.EmbedBuilder();
  embed.setTitle("KLESS HIVE ROSTER");
  embed.setColor("Random");
  embed.setFooter({
    text: "⚔️ Assassin (Melee), 🏹 Assassin (Ranged), 🟠 Bruiser, 🔵 DPS: Blunderbuss/Ice, 🟡 DPS: Fire staff, 🟢 Healer, 🟣 Support, 🦾 Tank",
  });
  for (let i = 0; i < columnData.length; i++) {
    const element = columnData[i];
    var players = [];
    await element.players.forEach((player, i) => {
      if (
        client.guilds.cache
          .get(process.env.GUILD)
          .members.cache.find(
            (user) =>
              user.nickname == player.name || user.displayName == player.name
          )
      ) {
        const user = client.guilds.cache
          .get(process.env.GUILD)
          .members.cache.find(
            (user) =>
              user.nickname == player.name || user.displayName == player.name
          );

        players.push(replaceRoleName(player.role) + " <@" + user.id + ">");
      } else {
        players.push(replaceRoleName(player.role) + " " + player.name);
      }
    });

    embed.addFields({
      name: element.name,
      value: players.join(" \n"),
      inline: true,
    });
    if ((i + 1) % 3 == 0) {
      embed.addFields({ name: "\u200b", value: "\u200b" });
    }
  }
  if (config.last_message) {
    client.channels.fetch(config.channel).then(async (channel) => {
      await channel.messages.fetch(config.last_message).then((msg) => {
        msg.edit({ embeds: [embed] });
      });
    });
  } else {
    await client.guilds.cache
      .get(process.env.GUILD)
      .channels.cache.get(config.channel)
      .send({ embeds: [embed] })
      .then((res) => {
        var newConfig = config;
        newConfig.last_message = res.id;
        fs.writeFile(
          "./config.json",
          JSON.stringify(newConfig, null, 2),
          (err) => {
            if (err) console.log(err);
          }
        );
      });
  }
}

// Function to read Excel data
async function readExcel(filePath) {
  return new Promise(async (resolve, reject) => {
    await axios
      .get(process.env.URL, {
        responseType: "arraybuffer",
      })
      .then((res) => {
        const workbook = xlsx.read(res.data, {
          type: "buffer",
        });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        resolve(xlsx.utils.sheet_to_json(worksheet));
      })
      .catch((err) => {
        console.log(err);
        reject();
      });
  });
}

//Generate Proper json
async function generateProperJSON(data) {
  var columnData = [];

  for (let j = 0; j < Object.keys(data[0]).length; j = j + 2) {
    const newGroup = {
      name: "",
      players: [],
    };
    newGroup.name = "Group" + (j + 2) / 2;
    for await (const element of data) {
      newGroup.players.push({
        name: Object.values(element)[j],
        role: Object.values(element)[j + 1],
      });
    }
    columnData.push(newGroup);
  }
  return columnData;
}

////////REPLACE ROLENAME

function replaceRoleName(role) {
  return role
    .replace("Assassin (Melee)", "⚔️")
    .replace("Assassin (Ranged)", "🏹")
    .replace("Bruiser", "🟠")
    .replace("DPS: Blunderbuss/Ice", "🔵")
    .replace("DPS: Fire staff", "🟡")
    .replace("Healer", "🟢")
    .replace("Support", "🟣")
    .replace("Tank", "🦾");
}

// Command to import and send Excel data to Discord
client.on("messageCreate", async (message) => {
  /////////import command
  if (message.content === "n!import") {
    if (!message.member.permissionsIn(message.channel).has("Administrator"))
      return message.channel.send("You are not an admin");
    const data = await readExcel("./test.xlsx");
    const columnData = await generateProperJSON(data);
    const embed = new Discord.EmbedBuilder();
    embed.setTitle("KLESS HIVE ROSTER");
    embed.setColor("Random");
    embed.setFooter({
      text: "⚔️ Assassin (Melee), 🏹 Assassin (Ranged), 🟠 Bruiser, 🔵 DPS: Blunderbuss/Ice, 🟡 DPS: Fire staff, 🟢 Healer, 🟣 Support, 🦾 Tank",
    });
    for (let i = 0; i < columnData.length; i++) {
      const element = columnData[i];
      var players = [];
      await element.players.forEach((player, i) => {
        if (
          client.guilds.cache
            .get(process.env.GUILD)
            .members.cache.find(
              (user) =>
                user.nickname == player.name || user.displayName == player.name
            )
        ) {
          const user = client.guilds.cache
            .get(process.env.GUILD)
            .members.cache.find(
              (user) =>
                user.nickname == player.name || user.displayName == player.name
            );
          players.push(replaceRoleName(player.role) + " <@" + user.id + ">");
        } else {
          players.push(replaceRoleName(player.role) + " " + player.name);
        }
      });

      embed.addFields({
        name: element.name,
        value: players.join(" \n"),
        inline: true,
      });
      if ((i + 1) % 3 == 0) {
        embed.addFields({ name: "\u200b", value: "\u200b" });
      }
    }
    await message.channel.send({ embeds: [embed] });
  }

  ///////////////////////// update command//////////////////////
  if (message.content == "n!updatechannel" && !message.author.bot) {
    if (!message.member.permissionsIn(message.channel).has("Administrator"))
      return message.channel.send("You are not an admin");
    message.channel.send("Done").then((msg) => {
      setTimeout(() => {
        msg.delete();
        message.delete();
      }, 5000);
    });
    var newConfig = config;
    newConfig.channel = message.channel.id;
    newConfig.last_message = "";
    fs.writeFile(
      "./config.json",
      JSON.stringify(newConfig, null, 2),
      (error) => {
        if (err) {
          console.log("error while saving");
        } else {
          console.log("saved!");
        }
      }
    );
  }

  //////////////////////stop command.//////////////////////////
  if (message.content == "n!stopupdate") {
    if (!message.member.permissionsIn(message.channel).has("Administrator"))
      return message.channel.send("You are not an admin");
    var newConfig = config;
    newConfig.channel = "";
    newConfig.last_message = "";
    fs.writeFile(
      "./config.json",
      JSON.stringify(newConfig, null, 2),
      (error) => {
        if (err) {
          console.log("error while saving");
        } else {
          console.log("saved!");
          message.channel.send("Stopped updating");
        }
      }
    );
  }

  ////////////////////help command ////////////////////////////
  if (message.content == "n!help") {
    const embed = new Discord.EmbedBuilder();
    embed.setTitle("My commands are:");
    embed.addFields({ name: "n!stopupdate", value: "Stop updating roles" });
    embed.addFields({
      name: "n!updatechannel",
      value: "Set channel to send updates",
    });
    embed.addFields({ name: "n!import", value: "Sends you role embed" });
    message.channel.send({ embeds: [embed] });
  }
});
