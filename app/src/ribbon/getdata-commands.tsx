import { FileText, Globe, Table2 } from "lucide-react"
import { canReadClipboardImage } from "@/pdf/picture"
import { forgetRecentSources, message, recentSources, type RecentSource } from "@/sheet/get-data"
import type { Command, RibbonEnv } from "./registry"

/**
 * Data › Get & Transform Data: Get Data's menu, From Web, From Table/Range, Recent Sources and Existing Connections
 * (From Text/CSV is in commands.tsx with the other file commands).
 */

const sourceLabel = (s: RecentSource) => (s.kind === "web" ? s.url : s.kind === "pdf" ? s.name : `Table ${s.name}`)

function openRecent(env: RibbonEnv, s: RecentSource) {
  env.c.getData
    .openRecent(s)
    .catch((e: unknown) => env.c.shell.alert(`Cells couldn't read ${sourceLabel(s)}. ${message(e)}`))
}

function RecentPanel({ env, close }: { env: RibbonEnv; close: () => void }) {
  const list = recentSources()
  return (
    <div className="recent-sources" role="menu" aria-label="Recent Sources">
      <div className="recent-sources-heading">Recent Sources</div>
      {list.length ? (
        list.map((s) => (
          <button
            key={sourceLabel(s)}
            type="button"
            role="menuitem"
            className="recent-sources-item"
            title={sourceLabel(s)}
            onClick={() => {
              close()
              openRecent(env, s)
            }}
          >
            {s.kind === "web" ? (
              <Globe size={16} aria-hidden />
            ) : s.kind === "pdf" ? (
              <FileText size={16} aria-hidden />
            ) : (
              <Table2 size={16} aria-hidden />
            )}
            <span>{sourceLabel(s)}</span>
          </button>
        ))
      ) : (
        <p className="recent-sources-empty">Sources you get data from show here.</p>
      )}
      {list.length ? (
        <button
          type="button"
          role="menuitem"
          className="recent-sources-item recent-sources-clear"
          onClick={() => {
            forgetRecentSources()
            close()
          }}
        >
          Clear Recent Sources
        </button>
      ) : null}
    </div>
  )
}

export function getDataCommands(): Command[] {
  return [
    {
      id: "PowerQueryGetTransformDataMenuSubscription19",
      label: "Get Data",
      description: "Bring in data from a file, the web or a table, and shape it in the Query Editor before it loads.",
      keytip: "PN",
      menu: [
        { heading: "From File" },
        "GetPowerQueryGetTransformDataFromText",
        "PowerQueryFromPdf",
        { heading: "From Other Sources" },
        "GetPowerQueryDataFromWeb",
        "GetPowerQueryGetTransformDataFromTable",
        "-",
        "PowerQueryRecentSourcesItem",
        "GetTransformExistingConnections",
        "ShowQueriesAndConnections",
      ],
    },
    {
      id: "PowerQueryFromPdf",
      label: "From PDF",
      description: "Get the tables in a PDF: Cells finds them on its pages, and you pick the ones to load.",
      keytip: "P",
      run: (env) => void env.files?.importPdf(),
      enabled: (env) => Boolean(env.files),
    },
    {
      id: "GetPowerQueryDataFromWeb",
      label: "From Web",
      description: "Get data from a CSV or JSON file on the web, or from a table on a web page.",
      keytip: "W",
      run: (env) => env.dialog("fromWeb"),
    },
    {
      id: "GetPowerQueryGetTransformDataFromTable",
      label: "From Table/Range",
      description: "Make a query of the selected table or range, to filter, sort and reshape it into a new table.",
      keytip: "T",
      run: (env) => void env.c.getData.fromTable(),
    },
    {
      id: "DataFromPicture",
      label: "From Picture",
      description:
        "Turn a picture of a table, such as a photo or a screenshot, into cells: Cells reads it, and you check it before it goes in.",
      keytip: "PI",
      menu: ["DataFromPictureFile", "DataFromPictureClipboard"],
      enabled: (env) => Boolean(env.files),
    },
    {
      id: "DataFromPictureFile",
      label: "Picture From File...",
      description: "Pick a picture of a table to read into cells.",
      keytip: "F",
      run: (env) => void env.files?.importPicture(),
      enabled: (env) => Boolean(env.files),
    },
    {
      id: "DataFromPictureClipboard",
      label: "Picture From Clipboard",
      description: "Read the picture you copied, such as a screenshot of a table, into cells.",
      keytip: "C",
      run: (env) => void env.files?.pastePicture(),
      enabled: (env) => Boolean(env.files) && canReadClipboardImage(),
    },
    {
      id: "PowerQueryRecentSources",
      label: "Recent Sources",
      description: "Get data again from a source you used before.",
      keytip: "PR",
      menu: [{ panel: (env, close) => <RecentPanel env={env} close={close} /> }],
    },
    {
      id: "PowerQueryRecentSourcesItem",
      label: "Recent Sources",
      description: "Get data again from a source you used before.",
      menu: [{ panel: (env, close) => <RecentPanel env={env} close={close} /> }],
    },
    {
      id: "GetTransformExistingConnections",
      label: "Existing Connections",
      description: "See the workbook's connections and load one, or go to the data it fills.",
      keytip: "X",
      run: (env) => env.dialog("existingConnections"),
    },
    {
      id: "AsppRelationships",
      label: "Relationships",
      description: "See how the tables in the Data Model are related, so a PivotTable can combine them.",
      keytip: "RN",
      run: (env) => env.dialog("relationships"),
    },
    {
      id: "DataModelManage",
      label: "Manage Data Model",
      description: "See the Data Model's tables, their data, and a diagram of how they relate.",
      keytip: "M",
      run: (env) => env.dialog("dataModel"),
    },
  ]
}
