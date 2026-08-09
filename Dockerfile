FROM node:24-alpine AS build
WORKDIR /app
# The lock file, and `npm ci` rather than `npm install`: CI verifies the tree the
# lock pins, so resolving dependencies again here is how the image ends up
# running something no test ever ran. tsconfig.build.json too — it is what
# `npm run build` points tsc at, and the file that excludes the tests from dist.
COPY package.json package-lock.json tsconfig.json tsconfig.build.json ./
RUN npm ci
COPY src ./src
RUN npm run build

FROM node:24-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force
COPY --from=build /app/dist ./dist
# The fonts, which the PDF renderer reads at runtime. Not a build-time asset:
# PDF's built-in fonts cover Latin-1, so without these a Korean document renders
# as blank space. `dist/write/pdf.js` resolves them at ../../assets/fonts/.
COPY assets ./assets
# Nothing here writes to disk or binds a privileged port, so there is nothing
# root buys — and this process exists to parse bytes chosen by a model.
USER node
EXPOSE 3000
# exec form: node is PID 1 so SIGTERM reaches it on a rolling deploy
CMD ["node", "dist/server.js"]
