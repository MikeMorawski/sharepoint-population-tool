// Gets rid of code splitting chunks to provide a single entry point for parent project
module.exports = {
    webpack: {
        configure: (webpackConfig, { env, paths }) => {
            // Easier CSS reference
            const miniCssExtractPlugin = webpackConfig.plugins.find((p) => p.constructor.name === "MiniCssExtractPlugin");
            if (miniCssExtractPlugin) {
                miniCssExtractPlugin.options.filename = "static/css/[name].css"
            }

            // Disable chunking
            webpackConfig.optimization.runtimeChunk = false;
            webpackConfig.optimization.splitChunks = { chunks(chunk) { return false } }

            // Filename adjust
            webpackConfig.output.filename = './static/js/bundle.js';

            return webpackConfig;
        }
    },
}