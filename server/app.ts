import App from "./app.server"
import { FilesController } from "./controller/files.controller"
import dotenv from "dotenv"

dotenv.config()

const port = process.env.PORT || 8080

const app = new App([new FilesController()], port as number)

app.listen()
