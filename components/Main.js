import { Headline } from "../components/Headline";
import { Links } from "../components/Links";
import classes from "./Main.module.css";

export function Main(props) {
  return (
    <main className={classes.main}>
      <Headline page={props.page}>
        <code className={classes.code}>pages/{props.page}</code>
      </Headline>
      <Links />
    </main>
  );
}
