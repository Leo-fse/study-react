import { Headline } from "../components/Headline";
import { Links } from "../components/Links";
import styles from "../styles/Home.module.css";

export function Main(props) {
  return (
    <main className={styles.main}>
      <Headline page={props.page}>
        <code className={styles.code}>pages/{props.page}</code>
      </Headline>
      <Links />
    </main>
  );
}
