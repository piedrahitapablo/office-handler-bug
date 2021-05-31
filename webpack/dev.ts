import * as OfficeDevCerts from 'office-addin-dev-certs';
import * as webpack from 'webpack';
import { merge } from 'webpack-merge';

import common from './common';

const config = async (): Promise<webpack.Configuration> => {
  const server = await OfficeDevCerts.getHttpsServerOptions();

  return merge(common, {
    mode: 'development',
    devtool: 'inline-source-map',
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore: Needed because devServer is not present on webpack config type
    devServer: {
      hot: true,
      https: server,
      overlay: true,
      port: 3000,
    },
    plugins: [new webpack.HotModuleReplacementPlugin()],
  });
};

export default config;
