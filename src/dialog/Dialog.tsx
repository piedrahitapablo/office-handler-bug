import React, { useEffect, useState } from "react";

const Dialog: React.FC = () => {
  const [ready, setReady] = useState(false);
  const [message, setMessage] = useState<string | null>(null);

  useEffect(() => {
    Office.onReady()
      .then(
        () =>
          new Promise<void>((resolve, reject) => {
            Office.context.ui.addHandlerAsync(
              Office.EventType.DialogParentMessageReceived,
              (result) => setMessage(result.message),
              {},
              (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                  reject(result.error);
                  return;
                }

                resolve();
              }
            );
          })
      )
      .then(() => {
        Office.context.ui.messageParent("ready");
        setReady(true);
      });
  }, []);

  if (!ready) {
    return <>Loading...</>;
  }

  if (message == null) {
    return <>No message</>;
  }

  return <>{message}</>;
};

export default Dialog;
