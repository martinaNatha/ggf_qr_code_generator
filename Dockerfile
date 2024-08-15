FROM node
WORKDIR /usr/scr/app/
COPY package*.json ./
RUN npm install

COPY . .

CMD [ "npm", "start" ]