import express, { Request, Response } from "express";
import expressWs, { Application } from "express-ws";
import * as ws from "ws";
import axios, { AxiosError } from "axios";
import ExcelJS from "exceljs";

const app: Application = expressWs(express(), undefined, {
  wsOptions: { clientTracking: true },
}).app;

const port = Number(process.env.PORT ?? 3001);

let wsConnections: Record<string, ws> = {};

app.get("/missions", async (req, res) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Irène";
  workbook.lastModifiedBy = "Irène";
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheet = workbook.addWorksheet("My Sheet");

  let numberOfSC2SMissions = 0;

  sheet.columns = [
    { header: "Nom de l'annonce", key: "name", width: 32 },
    { header: "Organisme", key: "organism", width: 32 },
    { header: "Date", key: "date", width: 32 },
    { header: "Lieu de la mission (Ville)", key: "place", width: 32 },
    { header: "Lieu de la mission (Code postal)", key: "place", width: 32 },
    { header: "Contact", key: "contact", width: 32 },
    { header: "Téléphone", key: "phone", width: 32 },
    { header: "Email", key: "email", width: 32 },
    { header: "Lien vers l'annonce", key: "link", width: 32 },
    { header: "Publics bénéficiaires", key: "public", width: 32 },
  ];

  let missionsCount = 0;

  if (req.query.wsClientId) {
    wsConnections[req.query.wsClientId.toString()].send(
      "0 missions analysées."
    );
  }

  try {
    const response = await axios.get(
      `https://www.service-civique.gouv.fr/api/api/rest/missions?statusList%5B%5D=published&publicBeneficiaries=${req.query.publicBeneficiaries}&orderByField=publishDate&orderByDirection=DESC&first=${req.query.first}`
    );

    for (let i = 0; i < response.data.edges.length; i++) {
      const row = sheet.getRow(i + 2 - numberOfSC2SMissions);

      const edge = response.data.edges[i];

      if (
        (!edge.node.title.toLowerCase().includes("sc2s") ||
          req.query.excludeSC2S === "false") &&
        edge.node.status === "published"
      ) {
        row.getCell(1).value = edge.node.title;
        row.getCell(2).value = edge.node.organization.name;
        row.getCell(3).value = edge.node.startDate;
        row.getCell(4).value = edge.node.interventionPlace.city;
        row.getCell(5).value = edge.node.interventionPlace.zip;

        const missionResponse = await axios.get(
          `https://www.service-civique.gouv.fr/api/api/rest/missions/${edge.node.id}`
        );

        row.getCell(6).value =
          missionResponse.data.contact?.firstName +
          " " +
          missionResponse.data.contact?.lastName +
          " " +
          (missionResponse.data.contact?.function
            ? `(fonction: ${missionResponse.data.contact?.function})`
            : "");
        row.getCell(7).value = missionResponse.data.contact?.telephone;
        row.getCell(8).value = missionResponse.data.contact?.email;
        row.getCell(
          9
        ).value = `https://www.service-civique.gouv.fr/trouver-ma-mission/${edge.node.slug}-${edge.node.id}`;

        row.getCell(10).value =
          missionResponse.data.publicBeneficiaries.join(",");

        missionsCount++;
        if (req.query.wsClientId) {
          wsConnections[req.query.wsClientId.toString()].send(
            `${missionsCount} missions analysées.`
          );

          console.log(`Mission ${edge.node.id} added`);
        }
      } else {
        numberOfSC2SMissions++;
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();

    res.write(buffer);
    res.end();
  } catch (error) {
    const axiosError = error as AxiosError;
    console.error("error", axiosError);
    res
      .status(Number(axiosError.code))
      .send(`${axiosError.code} ${axiosError.message}`);
    res.end();
  }
});

app.ws("/ws", (ws, req) => {
  if (req.query.wsClientId) {
    wsConnections[req.query.wsClientId.toString()] = ws;
  }
});

app.get("/", (req, res) => {
  res.json("coucou");
});

app.listen(port, function () {
  console.log(`App is listening on port ${port} !`);
});
