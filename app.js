const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
const webpack = require("webpack");
const webPackConfig = require("./webpack.config");
const webPackDevMiddleware = require("webpack-dev-middleware");
const cors = require("cors");

const app = express();
// const compiler = webpack(webPackConfig);

app.use(bodyParser.json());
app.use(cors());

// app.use(
//   webPackDevMiddleware(compiler, {
//     publicPath: webPackConfig.output.publicPath,
//   })
// );
// app.use(express.static(path.join(__dirname, "src", "taskpane")));

app.use(express.static(path.join(__dirname, "src", "client")));
const router = require("./src/routes/index");
const files = require("./src/routes/getFilesRoute");

app.use("/", router);
app.use("/", files);
app.listen(3001);
console.log("server is listening on port 3001");
