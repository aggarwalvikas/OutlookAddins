import * as React from "react";

interface MessageDetails {
  subject: string;
  senderName: string;
  senderAddress: string;
}

const getMessageDetails = (): MessageDetails | null => {
  const item = Office.context.mailbox?.item as Office.MessageRead | undefined;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    return null;
  }

  const sender = item.from;
  return {
    subject: item.subject || "(No subject)",
    senderName: sender?.displayName || "Unknown sender",
    senderAddress: sender?.emailAddress || "Email address unavailable",
  };
};

const App: React.FC = () => {
  const [details, setDetails] = React.useState<MessageDetails | null>(null);

  React.useEffect(() => {
    setDetails(getMessageDetails());
  }, []);

  return (
    <main style={styles.page}>
      <h1 style={styles.heading}>Outlook Mail Viewer</h1>
      <p style={styles.intro}>Details for the message currently open in Outlook.</p>
      {details ? (
        <dl style={styles.details}>
          <div style={styles.detail}>
            <dt style={styles.label}>Subject</dt>
            <dd style={styles.value}>{details.subject}</dd>
          </div>
          <div style={styles.detail}>
            <dt style={styles.label}>Sender</dt>
            <dd style={styles.value}>
              {details.senderName}
              <br />
              {details.senderAddress}
            </dd>
          </div>
        </dl>
      ) : (
        <p style={styles.message}>Open an email message to view its subject and sender.</p>
      )}
    </main>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  page: {
    minHeight: "100vh",
    boxSizing: "border-box",
    padding: "24px 20px",
    backgroundColor: "#f5f5f5",
    color: "#242424",
    fontFamily: "Segoe UI, Arial, sans-serif",
  },
  heading: {
    margin: 0,
    color: "#0f6cbd",
    fontSize: "24px",
    fontWeight: 600,
  },
  intro: {
    margin: "8px 0 24px",
    color: "#616161",
    fontSize: "14px",
  },
  details: {
    display: "block",
    margin: 0,
  },
  detail: {
    margin: "0 0 16px",
    padding: "14px 16px",
    backgroundColor: "#ffffff",
    border: "1px solid #d1d1d1",
    borderRadius: "4px",
  },
  label: {
    display: "block",
    marginBottom: "6px",
    color: "#616161",
    fontSize: "12px",
    fontWeight: 600,
    textTransform: "uppercase",
  },
  value: {
    display: "block",
    overflowWrap: "break-word",
    fontSize: "16px",
  },
  message: {
    margin: 0,
    padding: "14px 16px",
    backgroundColor: "#ffffff",
    border: "1px solid #d1d1d1",
    borderRadius: "4px",
    color: "#616161",
    lineHeight: 1.5,
  },
};

export default App;
