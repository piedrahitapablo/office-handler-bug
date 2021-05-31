import React, { useEffect, useState } from "react";

const Taskpane: React.FC = () => {
  const [ready, setReady] = useState(false);
  const [dialog, setDialog] = useState<Office.Dialog | null>(null);

  useEffect(() => {
    Office.onReady().then(() => setReady(true));
  });

  const handleShowDialogWithRedirectClick = (): void => {
    const url =
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
      "?client_id=<PUT_YOUR_CLIENT_ID>" +
      "&response_type=token" +
      "&redirect_uri=https%3A%2F%2Flocalhost%3A3000%2Fdialog.html" +
      "&scope=email+profile+openid" +
      "&access_type=online" +
      "&prompt=consent";
    const options = {
      width: 20,
      height: 40,
      promptBeforeOpen: true,
      displayInIframe: false,
    };

    Office.context.ui.displayDialogAsync(url, options, (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        if ("message" in args && args.message === "ready") {
          setDialog(dialog);
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        setDialog(null);
      });
    });
  };

  const handleShowDialogClick = () => {
    const url = "https://localhost:3000/dialog.html";
    const options = {
      width: 20,
      height: 40,
      promptBeforeOpen: true,
      displayInIframe: false,
    };

    Office.context.ui.displayDialogAsync(url, options, (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        if ("message" in args && args.message === "ready") {
          setDialog(dialog);
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        setDialog(null);
      });
    });
  }

  const handleSendMessageClick = (): void => {
    if (dialog == null) {
      throw new Error("dialog is null");
    }

    dialog.messageChild("hi from the taskpane :)");
  };

  if (!ready) {
    return <>Loading...</>;
  }

  return (
    <div>
      {dialog == null && (
        <div>
        <button type="button" onClick={handleShowDialogClick}>
          Show dialog
        </button>
        <button type="button" onClick={handleShowDialogWithRedirectClick}>
          Show dialog with redirect
        </button>
        </div>
      )}

      {dialog != null && (
        <button type="button" onClick={handleSendMessageClick}>
          Send message to dialog
        </button>
      )}
    </div>
  );
};

export default Taskpane;
