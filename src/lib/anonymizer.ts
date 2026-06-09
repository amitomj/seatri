export function deduplicateEntities(entities: PIIEntity[]): PIIEntity[] {
  const uniqueEntities: PIIEntity[] = [];
  const seenIds = new Set<string>();
  
  entities.forEach(entity => {
    if (!seenIds.has(entity.id)) {
      uniqueEntities.push(entity);
      seenIds.add(entity.id);
    } else {
      // If duplicate ID found, give it a new one to prevent React key errors
      const newEntity = { ...entity, id: generateId() };
      uniqueEntities.push(newEntity);
      seenIds.add(newEntity.id);
    }
  });
  
  return uniqueEntities;
}

export function generateId(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return Math.random().toString(36).substring(2, 15) + Date.now().toString(36);
}

export type PIIType = string; // More flexible type system

export interface PIIEntity {
  id: string;
  original: string;
  type: PIIType;
  pseudonym: string;
  enabled: boolean;
  selected?: boolean;
  ignored?: boolean;
  treated?: boolean;
  groupId?: string;
  fileIds?: string[]; // IDs of files where this entity was found
  context?: string; // Word immediately preceding the entity
  contextBefore?: string; // Snippet before
  contextAfter?: string; // Snippet after
  contextSnippet?: string; // 2 words before + entity + 2 words after
  reviewed?: boolean; // For ambiguity review
  score?: number; // Confidence score
  page?: number; // Page number where it was found
  y?: number; // Vertical position (0 to 1, or absolute)
  x?: number; // Horizontal position
  isNewlyIdentified?: boolean; // Flag to highlight newly found entities during re-analysis
}

export const PII_COLORS: Record<string, { bg: [number, number, number], text: [number, number, number], hex: string, textHex: string }> = {
  NOME: { bg: [1, 0.9, 0], text: [0, 0, 0], hex: '#FFD700', textHex: '#000000' },      // Ouro
  LOCAL: { bg: [0.2, 0.8, 0.2], text: [1, 1, 1], hex: '#32CD32', textHex: '#FFFFFF' }, // Verde Lima
  PHONE: { bg: [0.1, 0.5, 0.9], text: [1, 1, 1], hex: '#1E90FF', textHex: '#FFFFFF' }, // Azul Dodger
  NIF: { bg: [1, 0.5, 0], text: [1, 1, 1], hex: '#FF8C00', textHex: '#FFFFFF' },      // Laranja Escuro
  CC: { bg: [0.9, 0.1, 0.5], text: [1, 1, 1], hex: '#DC143C', textHex: '#FFFFFF' },    // Carmesim
  PASSPORT: { bg: [0.6, 0.2, 0.8], text: [1, 1, 1], hex: '#9932CC', textHex: '#FFFFFF' }, // Orquídea Escura
  EMAIL: { bg: [0.4, 0.4, 0.4], text: [1, 1, 1], hex: '#696969', textHex: '#FFFFFF' }, // Cinza Escuro
  IBAN: { bg: [0.7, 0.6, 0.1], text: [1, 1, 1], hex: '#B8860B', textHex: '#FFFFFF' },  // Ouro Velho
  AUTOR: { bg: [0, 0, 0], text: [1, 1, 1], hex: '#000000', textHex: '#FFFFFF' },      // Preto
  JUIZ: { bg: [0, 0, 0.5], text: [1, 1, 1], hex: '#000080', textHex: '#FFFFFF' },     // Azul Marinho
  MATRICULA: { bg: [0.8, 0.8, 0], text: [0, 0, 0], hex: '#CCCC00', textHex: '#000000' }, // Amarelo Escuro
  ADVOGADO: { bg: [0.5, 0.5, 0.5], text: [1, 1, 1], hex: '#808080', textHex: '#FFFFFF' }, // Cinza
  COLETIVA: { bg: [0.5, 0.2, 0.1], text: [1, 1, 1], hex: '#8B4513', textHex: '#FFFFFF' }, // Marrom (SaddleBrown)
  HEADER: { bg: [0.9, 0.9, 0.9], text: [0.4, 0.4, 0.4], hex: '#E5E7EB', textHex: '#4B5563' }, // Cinza Claro
};

// Dicionário de nomes comuns em Portugal para aumentar a precisão
const COMMON_PT_FIRST_NAMES = new Set([
  'Maria', 'Ana', 'João', 'José', 'Francisco', 'António', 'Manuel', 'Carlos', 'Pedro', 'Paulo',
  'Luís', 'Miguel', 'Ricardo', 'Nuno', 'Rui', 'Vítor', 'Jorge', 'Joaquim', 'Fernando', 'Luísa',
  'Isabel', 'Teresa', 'Margarida', 'Helena', 'Cristina', 'Rita', 'Sofia', 'Inês', 'Mariana', 'Beatriz',
  'Tiago', 'Gonçalo', 'Diogo', 'André', 'Filipe', 'Duarte', 'Afonso', 'Rodrigo', 'Martim', 'Lourenço',
  'Guilherme', 'Gabriel', 'Mateus', 'Lucas', 'Daniel', 'David', 'Samuel', 'Rafael', 'Hugo', 'Bruno',
  'Sérgio', 'Marco', 'Alexandre', 'Nelson', 'Eduardo', 'Artur', 'Mário', 'Álvaro', 'Augusto', 'Alfredo',
  'Leonor', 'Matilde', 'Carolina', 'Alice', 'Benedita', 'Francisca', 'Clara', 'Laura', 'Vera', 'Marta',
  'Sara', 'Diana', 'Patrícia', 'Sílvia', 'Sandra', 'Paula', 'Carla', 'Mónica', 'Cláudia', 'Sónia',
  'Catarina', 'Daniela', 'Andreia', 'Joana', 'Filipa', 'Alexandra', 'Telma', 'Tânia', 'Vanessa', 'Rute',
  'Marcelo', 'Bernardo', 'Gustavo', 'Henrique', 'Leonardo', 'Salvador', 'Tomás', 'Vicente', 'Valter',
  'Cláudio', 'Dinis', 'Emanuel', 'Fábio', 'Gil', 'Hélder', 'Igor', 'Jaime', 'Kevin', 'Leandro',
  'Mauro', 'Nicolau', 'Otávio', 'Patrício', 'Quelvin', 'Renato', 'Simão', 'Tadeu', 'Urbano', 'Vasco',
  'Wilson', 'Xavier', 'Yuri', 'Zacarias', 'Abel', 'Adelaide', 'Adolfo', 'Adriana', 'Adriano', 'Agostinho',
  'Albano', 'Alberto', 'Albina', 'Alda', 'Amadeu', 'Amélia', 'Américo', 'Amílcar', 'Anabela', 'Anacleto',
  'Antero', 'Apolinário', 'Arnaldo', 'Arsenio', 'Ascensão', 'Assunção', 'Áurea', 'Aurélio', 'Aurora',
  'Bárbara', 'Basílio', 'Belmiro', 'Bento', 'Berta', 'Bia', 'Branca', 'Brites', 'Caetano', 'Cândida',
  'Cândido', 'Carina', 'Carminda', 'Carmo', 'Casimiro', 'Cecília', 'Celeste', 'Célia', 'Celina', 'Celso',
  'César', 'Cidália', 'Ciro', 'Constantino', 'Cosme', 'Custódio', 'Dália', 'Dalila', 'Damaso', 'Damião',
  'Dário', 'Delfim', 'Delfina', 'Deolinda', 'Desidério', 'Diamantino', 'Dina', 'Dionísio', 'Domingos',
  'Dora', 'Dores', 'Doroteia', 'Edgar', 'Edite', 'Edmundo', 'Elisa', 'Elisabete', 'Elmano', 'Elsa',
  'Elvira', 'Elza', 'Emília', 'Emílio', 'Ercília', 'Ernesto', 'Esmeralda', 'Estefânia', 'Estêvão',
  'Estrela', 'Etelvina', 'Eugénia', 'Eugénio', 'Eulália', 'Eusébio', 'Eva', 'Evaristo', 'Evelina',
  'Faustino', 'Fausto', 'Felicidade', 'Feliciano', 'Felisberto', 'Felismina', 'Fidélio', 'Filomena',
  'Firmino', 'Flávia', 'Flávio', 'Florbela', 'Florença', 'Florinda', 'Fortunato', 'Frederico', 'Gaspar',
  'Genciana', 'Genoveva', 'Geraldo', 'Gertrudes', 'Gisela', 'Glória', 'Gracinda', 'Graça', 'Gregório',
  'Gualter', 'Guarani', 'Guendolina', 'Guiomar', 'Haroldo', 'Heitor', 'Helder', 'Hélder', 'Heliodoro',
  'Hélio', 'Henriqueta', 'Heraldo', 'Herberto', 'Herculano', 'Hermenegildo', 'Hermínia', 'Hermínio',
  'Higino', 'Hilario', 'Hilário', 'Hipólito', 'Horácio', 'Hortênsia', 'Ilda', 'Ilídio', 'Inácio',
  'Iolanda', 'Irene', 'Irineu', 'Íris', 'Isidoro', 'Isidro', 'Isilda', 'Ismael', 'Iva', 'Ivan',
  'Ivone', 'Izidoro', 'Jacinta', 'Jacinto', 'Jandira', 'Januário', 'Jesuína', 'Jesuíno', 'Jesus',
  'Job', 'Joel', 'Jonas', 'Jordão', 'Josué', 'Jovita', 'Judite', 'Júlia', 'Juliana', 'Juliano',
  'Julieta', 'Júlio', 'Justina', 'Justino', 'Juvenal', 'Lara', 'Lázaro', 'Lélia', 'Lénia', 'Lia',
  'Libânio', 'Licínio', 'Lídia', 'Lidiana', 'Liete', 'Lígia', 'Lília', 'Liliana', 'Lina', 'Lino',
  'Lívia', 'Lopo', 'Lorena', 'Lourdes', 'Lucas', 'Lúcia', 'Luciana', 'Luciano', 'Lucília', 'Lucílio',
  'Lucinda', 'Lucindo', 'Lucrécia', 'Ludgero', 'Ludovina', 'Luís', 'Luísa', 'Lurdes', 'Luzia',
  'Madalena', 'Mafalda', 'Magda', 'Malvina', 'Manuela', 'Mara', 'Marcelina', 'Marcelino', 'Marcília',
  'Márcio', 'Margarida', 'Maria', 'Mariana', 'Marília', 'Marina', 'Marinho', 'Mário', 'Marisa',
  'Marlene', 'Marta', 'Martim', 'Martinho', 'Mateus', 'Matilde', 'Maura', 'Maurício', 'Maximiliano',
  'Maximino', 'Melinda', 'Mélvin', 'Mercês', 'Messias', 'Micael', 'Micaela', 'Milton', 'Mílton',
  'Miriam', 'Modesto', 'Moisés', 'Mónica', 'Morgana', 'Muriel', 'Nádia', 'Nadir', 'Narciso',
  'Natália', 'Natalina', 'Natalino', 'Nazaré', 'Neide', 'Nélia', 'Nélio', 'Nélson', 'Nestor',
  'Neuza', 'Nicanor', 'Nícia', 'Nivaldo', 'Noé', 'Noel', 'Noémia', 'Norberto', 'Nuno', 'Octávio',
  'Odete', 'Ofélia', 'Olga', 'Olímpia', 'Olímpio', 'Olinda', 'Olindo', 'Olívia', 'Onofre', 'Orlandina',
  'Orlando', 'Oscar', 'Óscar', 'Osvaldo', 'Otília', 'Ouvídio', 'Palmira', 'Pancrácio', 'Pandora',
  'Pascoal', 'Pastor', 'Patrícia', 'Patrício', 'Paula', 'Paulina', 'Paulino', 'Paulo', 'Pedro',
  'Penélope', 'Perpétua', 'Piedade', 'Plácido', 'Policarpo', 'Prudêncio', 'Pura', 'Purificação',
  'Quitéria', 'Rachel', 'Raimundo', 'Raquel', 'Raul', 'Rebeca', 'Regina', 'Reinaldo', 'Remígio',
  'Renata', 'Renato', 'Ricardina', 'Ricardo', 'Rita', 'Rivaldo', 'Roberto', 'Rodolfo', 'Rodrigo',
  'Rogério', 'Rolando', 'Romeu', 'Rómulo', 'Rosa', 'Rosália', 'Rosalina', 'Rosário', 'Rúben',
  'Rui', 'Rute', 'Sabina', 'Salomão', 'Salomé', 'Salvador', 'Samuel', 'Sandra', 'Sandro', 'Santiago',
  'Sara', 'Sebastião', 'Selma', 'Serafim', 'Serena', 'Sérgio', 'Severino', 'Sílvia', 'Silvestre',
  'Silvia', 'Silvino', 'Simão', 'Sira', 'Siro', 'Sofia', 'Solange', 'Sónia', 'Soraia', 'Stela',
  'Susana', 'Tânia', 'Tasso', 'Tatiana', 'Telma', 'Telmo', 'Teodoro', 'Teófilo', 'Teresa', 'Tiago',
  'Tibério', 'Tito', 'Tomás', 'Tristão', 'Urbano', 'Valdemar', 'Valentim', 'Valentina', 'Valeriana',
  'Valeriano', 'Valério', 'Valter', 'Vanda', 'Vanessa', 'Vânia', 'Vasco', 'Venâncio', 'Vera',
  'Vicente', 'Victor', 'Vidal', 'Virgílio', 'Virgínia', 'Viriato', 'Vital', 'Vítor', 'Vitória',
  'Vitorino', 'Viviana', 'Xavier', 'Zélia', 'Zelinda', 'Zenaide', 'Zita', 'Zoe', 'Zulmira',
  'aabirah', 'aabaj', 'aaditi', 'aagambir', 'aaira', 'aahan', 'aaish', 'aahil', 'aakriti', 'aankit', 'aalia', 'aarav', 'aaliya', 'aaravpreet', 'aaliyah', 'aaron', 'aalyiah', 'aarón', 'aamna', 'aarush', 'aarica', 'aaryan', 'aarohi', 'aasaal', 'aarushi', 'aayam', 'aarvi', 'aayan', 'aashvi', 'aayun',
  'aatika', 'aayush', 'aayat', 'abd', 'aayushi', 'abdel', 'abby', 'abdelhadi', 'abbygaëlle', 'abdellah', 'abeedha', 'abdelrahman', 'abeeha', 'abderrahmane', 'abhinaya', 'abdoallah', 'abigael', 'abdoulaye', 'abigail', 'abdrahman', 'abimaela', 'abducadre', 'abrayene', 'abdul', 'abril', 'abdulai', 'abrish', 'abdulaí', 'acácia', 'abdullah', 'acakarein', 'abdulmuiz', 'achia', 'abdularahman', 'acricia', 'abdulwahab', 'açucena', 'abduramane', 'ada', 'abdurrahim', 'adalbmira', 'abel', 'adalgiza', 'abhay', 'adália', 'abhayjit', 'adama', 'abhinav', 'adan', 'abhinoor',
  'adanna', 'abhiyan', 'addis', 'abibo', 'adel', 'abiel', 'adelaide', 'abilio', 'adele', 'abílio', 'adélia', 'abir', 'adelina', 'abner', 'adelle', 'abou', 'adely', 'aboubacar', 'adelyna', 'abraão', 'adeya', 'abraar', 'adiane', 'abraham', 'adília', 'abraim', 'aditi', 'abrão', 'adja', 'abrar', 'adje', 'absalão', 'adla', 'abubacar', 'adriana', 'acácio', 'adriane', 'adab', 'adrianna', 'adalberto', 'adriela', 'adam', 'adriele', 'adama', 'adrielle', 'adame', 'adrielly', 'adão', 'adriely', 'adarsh',
  'adris', 'adaylton', 'adryelle', 'addison', 'adymara', 'áddison', 'afonsa', 'adeep', 'afonsina', 'adelin', 'áfrika', 'adelino', 'agashti', 'adélio', 'agata', 'adelmiro', 'ágata', 'adelson', 'agatha', 'adem', 'ágatha', 'aglaya', 'ademir', 'agnes', 'adenay', 'agnia', 'áder', 'águeda', 'adérito', 'aguinalda', 'ãdi', 'agustina', 'adiel', 'ahana', 'adil', 'ahinara', 'adilson', 'aicha', 'adnal', 'aida', 'adnan', 'aidin', 'adnav', 'aihra', 'adner', 'aíko', 'adney', 'aila', 'adolfo', 'ailine',
  'adon', 'ailla', 'adonai', 'aimara', 'adonay', 'aimee', 'adonias', 'aina', 'adónis', 'ainara', 'adriaan', 'ainhoa', 'adrian', 'aini', 'adrián', 'ainny', 'adriano', 'ainoa', 'adrick', 'ainoah', 'adriel', 'aisha', 'adrien', 'aísha', 'ádrien', 'aishleen', 'adryan', 'aishwarya', 'adryán', 'aissa', 'adryel', 'aissatá', 'adulai', 'aissatou', 'ady', 'aïssatou', 'aécio', 'aissatu', 'aedan', 'aitana', 'aeric', 'aixa', 'afeef', 'aiyane', 'affaan', 'aiza', 'afidjo', 'aizah', 'afonso', 'ajwa',
  'akari', 'akasha', 'akemi', 'akhil', 'akshara', 'akshat', 'akshaya', 'akshita', 'al', 'ala', 'alaa', 'alaba', 'aladin', 'aladino', 'alain', 'alair', 'alana', 'alan', 'alane', 'alani', 'alania', 'alanis', 'alanna', 'alano', 'alanya', 'alarico', 'alay', 'alaya', 'alayah', 'alayna', 'alba', 'alban', 'albana', 'albano', 'alberico', 'albert', 'alberta', 'albertha', 'alberti', 'albertina', 'albertino', 'alberto', 'albi', 'albina', 'albino', 'alcaide', 'alcântara', 'alceu', 'alci',
  'alcibiades', 'alcides', 'alcina', 'alcindo', 'alcino', 'alcion', 'alcione', 'alcir', 'alda', 'aldair', 'aldara', 'aldemira', 'aldemiro', 'aldenir', 'aldenora', 'alder', 'aldérico', 'aldina', 'aldino', 'aldir', 'aldo', 'aldomiro', 'aldora', 'aleandro', 'alec', 'alecsander', 'alef', 'alefe', 'alegria', 'aleida', 'alejandra', 'alejandro', 'aleksander', 'aleksandra', 'alen', 'alena', 'alencar', 'alendra', 'alene', 'aleni', 'alenia', 'alenice', 'alenir', 'alenis', 'alenka', 'alenquer', 'alentina', 'alento', 'alenza', 'aleonor',
  'aléscia', 'alesandra', 'alessandra', 'alessandro', 'alessia', 'aléssia', 'alessio', 'aléssio', 'alex', 'alexa', 'alexander', 'alexandra', 'alexandre', 'alexandrina', 'alexandrino', 'alexandro', 'alexia', 'aléxia', 'alexis', 'alexsander', 'alexsandra', 'alexsandre', 'alexsandro', 'aleya', 'alfaiate', 'alfaro', 'alfe', 'alfeu', 'alfia', 'alfie', 'alfio', 'alfons', 'alfonsa', 'alfonsina', 'alfonso', 'alfonz', 'alfred', 'alfreda', 'alfredo', 'alfrida', 'alfridi', 'alfrido', 'alga', 'algayer', 'algemira', 'algemiro', 'alhandra', 'ali', 'alia', 'alía',
  'aliah', 'aliana', 'aliane', 'alianna', 'aliano', 'aliar', 'alias', 'aliat', 'aliaz', 'aliba', 'alibe', 'alibi', 'alibo', 'alice', 'alicia', 'alícia', 'alicio', 'alício', 'alida', 'alídia', 'alie', 'alif', 'alife', 'alika', 'alike', 'aliki', 'aliko', 'alila', 'alim', 'alima', 'alime', 'alimi', 'alimo', 'alin', 'alina', 'alinda', 'aline', 'alini', 'alino', 'aliny', 'alípio', 'alique', 'alira', 'alis', 'alisa', 'alise', 'alisha', 'alisia', 'alísia', 'alison',
  'alisson', 'alister', 'alita', 'alito', 'aliu', 'alivia', 'alix', 'alixa', 'alixe', 'alixi', 'alixo', 'aliyah', 'aliyana', 'aliyane', 'aliz', 'aliza', 'alize', 'alizee', 'alizia', 'alízio', 'alizze', 'aljandira', 'aljandiro', 'aljara', 'aljare', 'aljari', 'aljaro', 'aljira', 'aljire', 'aljiri', 'aljiro', 'alju', 'aljure', 'aljuri', 'aljuro', 'alkis', 'all', 'alla', 'allah', 'allan', 'allana', 'allane', 'allani', 'allano', 'allany', 'allara', 'allare', 'allari', 'allaro',
  'adange', 'adanis', 'adanson', 'adão', 'add', 'adega', 'adela', 'adelária', 'adelino', 'adelo', 'adérito', 'aderneiro', 'adida', 'adler', 'admans', 'adolfíno', 'adolfo', 'adónis', 'adopel', 'adosindo', 'adrêgo', 'adriano', 'adrião', 'adrod', 'aduenes', 'aduleiro', 'advento', 'advincula', 'aerschot', 'aeugeneyndt', 'afé', 'afecto', 'afenar', 'affable', 'afilhado', 'afira', 'aflalo', 'aflição', 'aflitos', 'afogado', 'afoito', 'aforeira', 'afortunato', 'afra', 'afragoas', 'afreixio', 'afreixo', 'afresco'
].map(n => n.toLowerCase()));

const COMMON_PT_SURNAMES = new Set([
  'Silva', 'Santos', 'Ferreira', 'Pereira', 'Oliveira', 'Costa', 'Rodrigues', 'Martins', 'Jesus', 'Sousa',
  'Fernandes', 'Gonçalves', 'Gomes', 'Lopes', 'Marques', 'Alves', 'Almeida', 'Ribeiro', 'Pinto', 'Carvalho',
  'Teixeira', 'Moreira', 'Correia', 'Mendes', 'Nunes', 'Soares', 'Vieira', 'Monteiro', 'Cardoso', 'Rocha',
  'Raposo', 'Neves', 'Coelho', 'Cruz', 'Machado', 'Pires', 'Afonso', 'Dias', 'Duarte', 'Freitas',
  'Guerreiro', 'Henriques', 'Jorge', 'Leal', 'Leite', 'Lima', 'Maia', 'Melo', 'Miranda', 'Mota',
  'Nascimento', 'Neto', 'Pacheco', 'Paiva', 'Passos', 'Reis', 'Resende', 'Sá', 'Sampaio', 'Sanches',
  'Serrano', 'Simões', 'Tavares', 'Valente', 'Vaz', 'Vicente', 'Xavier', 'Brito', 'Cunha', 'Figueiredo',
  'Fonseca', 'Guerra', 'Laranjeira', 'Magalhães', 'Matos', 'Nogueira', 'Pinheiro', 'Queirós', 'Ramalho',
  'Sequeira', 'Valério', 'Abade', 'Abranches', 'Abrantes', 'Abreu', 'Agostinho', 'Aguiar', 'Aires',
  'Albuquerque', 'Alcântara', 'Aleixo', 'Alexandre', 'Almada', 'Almeida', 'Alonço', 'Altamirano',
  'Alvares', 'Alves', 'Alvim', 'Amado', 'Amaral', 'Amaro', 'Amorim', 'Andrade', 'Anes', 'Anjos',
  'Antunes', 'Aragão', 'Araújo', 'Areias', 'Arrais', 'Arruda', 'Assis', 'Assunção', 'Azevedo',
  'Bacelar', 'Badajoz', 'Baía', 'Bairros', 'Baldaia', 'Balsemão', 'Bandeira', 'Baptista', 'Barata',
  'Barbedo', 'Barbosa', 'Barcelos', 'Barreiros', 'Barreto', 'Barros', 'Barroso', 'Bastos', 'Batista',
  'Beça', 'Belchior', 'Belo', 'Beltrão', 'Bencatel', 'Bento', 'Bernardes', 'Bessa', 'Bettencourt',
  'Bezerra', 'Bicudo', 'Bivar', 'Boaventura', 'Botelho', 'Braga', 'Bragança', 'Brandão', 'Branco',
  'Brites', 'Brito', 'Brum', 'Bulhão', 'Cabral', 'Cabreira', 'Cacho', 'Cachopo', 'Cadaval', 'Caetano',
  'Caiado', 'Caires', 'Caldas', 'Caldeira', 'Calado', 'Camacho', 'Câmara', 'Camelo', 'Caminha',
  'Camões', 'Campelo', 'Campos', 'Canário', 'Candeias', 'Canedo', 'Cano', 'Cansado', 'Canto',
  'Capelo', 'Cardoso', 'Cardoso', 'Carlos', 'Carneiro', 'Carrascalão', 'Carreira', 'Carreiro',
  'Carrilho', 'Cartaxo', 'Carvalhal', 'Carvalho', 'Casado', 'Casal', 'Cascais', 'Castanheira',
  'Castelo', 'Castelo Branco', 'Castilho', 'Castro', 'Catela', 'Cavaco', 'Cavaleiro', 'Cerveira',
  'Chagas', 'Chaves', 'Cid', 'Cipriano', 'Claro', 'Clemente', 'Coelho', 'Coimbra', 'Colares',
  'Conceição', 'Conde', 'Cordeiro', 'Coronel', 'Correa', 'Correia', 'Corte-Real', 'Cortês', 'Costa',
  'Coutinho', 'Couto', 'Crespo', 'Cristóvão', 'Cruz', 'Cunha', 'Curado', 'Custódio', 'Damas',
  'Dâmaso', 'Dantas', 'Dantas', 'Delgado', 'Dinis', 'Diniz', 'Duarte', 'Eanes', 'Eiró', 'Encarnação',
  'Esteves', 'Evangelista', 'Fagundes', 'Faria', 'Farinha', 'Faro', 'Faustino', 'Feijó', 'Feio',
  'Felgueiras', 'Feliciano', 'Félix', 'Fernandes', 'Ferrão', 'Ferraz', 'Ferreira', 'Ferro', 'Fialho',
  'Fidalgo', 'Figueira', 'Figueiredo', 'Figueiró', 'Filipe', 'Fitas', 'Fogaça', 'Folque', 'Fonseca',
  'Fontes', 'Fortunato', 'Fraga', 'Fragoso', 'Fraisão', 'França', 'Franco', 'Freire', 'Freitas',
  'Frois', 'Furtado', 'Galvão', 'Gama', 'Gameiro', 'Garção', 'Garcia', 'Garrido', 'Gaspar', 'Gentil',
  'Gil', 'Girão', 'Godinho', 'Góis', 'Gomes', 'Gonçalves', 'Gouveia', 'Graça', 'Gramacho', 'Granjo',
  'Guedes', 'Guerra', 'Guerreiro', 'Guimarães', 'Gusmão', 'Henriques', 'Hipólito', 'Holbeche',
  'Homem', 'Horta', 'Iglésias', 'Inácio', 'Iria', 'Isidoro', 'Jacinto', 'Janes', 'Jardim', 'Jesus',
  'Joanes', 'Jorge', 'Jordão', 'Júlio', 'Junqueria', 'Lacerda', 'Ladeira', 'Lagos', 'Lamego',
  'Lameira', 'Lança', 'Landim', 'Laranjeira', 'Larcher', 'Leal', 'Leão', 'Leite', 'Leitão', 'Leme',
  'Lemos', 'Lencastre', 'Leonídio', 'Lima', 'Lira', 'Lisboa', 'Lobato', 'Lobo', 'Lopes', 'Lopo',
  'Loureiro', 'Lourenço', 'Lousada', 'Lucena', 'Luís', 'Luz', 'Macedo', 'Machado', 'Maciel',
  'Madureira', 'Magalhães', 'Magro', 'Maia', 'Mainato', 'Malaquias', 'Malheiro', 'Malta', 'Mamede',
  'Mâncio', 'Manteigas', 'Mântua', 'Manuel', 'Mariz', 'Marques', 'Martins', 'Mascarenhas', 'Mata',
  'Matos', 'Medeiros', 'Meira', 'Meireles', 'Melancia', 'Mello', 'Melo', 'Mendes', 'Mendonça',
  'Menezes', 'Mesquita', 'Mexia', 'Miranda', 'Moita', 'Moleiro', 'Moniz', 'Montenegro', 'Monteiro',
  'Morais', 'Moreira', 'Morgado', 'Mota', 'Moura', 'Mourão', 'Moutinho', 'Muniz', 'Murtinho',
  'Nascimento', 'Navarro', 'Naves', 'Negrão', 'Negreiros', 'Neto', 'Neves', 'Nobre', 'Nóbrega',
  'Nogueira', 'Nogueira', 'Noronha', 'Novaes', 'Novais', 'Nunes', 'Ó', 'Oliveira', 'Onofre',
  'Ornelas', 'Osório', 'Ourique', 'Outeiro', 'Pacheco', 'Padilha', 'Padrão', 'Paes', 'Paiva',
  'Paixão', 'Palha', 'Palma', 'Palmeira', 'Pamplona', 'Pantaleão', 'Panteleão', 'Paranhos', 'Pardo',
  'Paredes', 'Parreira', 'Passos', 'Pastana', 'Pato', 'Patrício', 'Paula', 'Paulino', 'Paz',
  'Peçanha', 'Pêra', 'Pedrosa', 'Pedroso', 'Peixoto', 'Pena', 'Penha', 'Penteado', 'Peralta',
  'Perdigão', 'Pereira', 'Peres', 'Pessoa', 'Pestana', 'Picanço', 'Picado', 'Pimenta', 'Pimentel',
  'Pinheiro', 'Pinho', 'Pinto', 'Pires', 'Pisco', 'Pissarra', 'Pita', 'Pizarro', 'Poças', 'Póvoas',
  'Pontes', 'Portela', 'Porto', 'Portugal', 'Prado', 'Prata', 'Prates', 'Prego', 'Preto', 'Pronto',
  'Proença', 'Prudêncio', 'Quaresma', 'Queirós', 'Queiroz', 'Quental', 'Quinteiro', 'Quintino',
  'Quirino', 'Rabelo', 'Radaquisto', 'Ramalho', 'Ramires', 'Ramos', 'Rangel', 'Raposo', 'Rato',
  'Real', 'Rebelo', 'Rebocho', 'Rego', 'Regueira', 'Reis', 'Relvas', 'Resende', 'Ribeiro', 'Rico',
  'Rios', 'Rocha', 'Rodrigues', 'Rolim', 'Romão', 'Rosa', 'Rosado', 'Rosário', 'Sá', 'Sabrosa',
  'Sacramento', 'Saldanha', 'Sales', 'Salgado', 'Salgueiro', 'Salomão', 'Salter', 'Salvado',
  'Sampaio', 'Sanches', 'Santana', 'Santarém', 'Santiago', 'Santos', 'Saraiva', 'Sardinha',
  'Sarmento', 'Seabra', 'Seco', 'Segurado', 'Seixas', 'Semedo', 'Sequeira', 'Serpa', 'Serrão',
  'Serrano', 'Sertório', 'Sesinando', 'Severo', 'Silva', 'Silveira', 'Silvestre', 'Simas', 'Simões',
  'Sintra', 'Siqueira', 'Soares', 'Sobral', 'Sobreira', 'Sobrinho', 'Solano', 'Sotto-Mayor', 'Sousa',
  'Souto', 'Souto-Maior', 'Spínola', 'Tavares', 'Taveira', 'Teixeira', 'Teles', 'Temudo', 'Teodoro',
  'Terra', 'Teves', 'Tojal', 'Toledo', 'Tolentino', 'Torres', 'Toscano', 'Toste', 'Trancoso',
  'Trigueiros', 'Trindade', 'Uchoa', 'Urbano', 'Valadares', 'Valadão', 'Valença', 'Valente',
  'Valentim', 'Vale', 'Valério', 'Valido', 'Varela', 'Vargas', 'Vasconcelos', 'Vasques', 'Vaz',
  'Veiga', 'Velasques', 'Veloso', 'Venda', 'Ventura', 'Vera-Cruz', 'Viana', 'Vicente', 'Vidal',
  'Viegas', 'Vieira', 'Vila-Lobos', 'Vilar', 'Vilhena', 'Vinhola', 'Vinhas', 'Viveiros', 'Xavier',
  'Ximenes', 'Zarco',
  'freitas', 'magalhães', 'henriques', 'lima', 'guerreiro', 'batista', 'pinheiro', 'faria', 'miranda', 'barros', 'morais', 'nogueira', 'esteves', 'anjos', 'baptista', 'campos', 'mota', 'andrade', 'brito', 'sá', 'nascimento', 'leite', 'abreu', 'borges', 'melo', 'vaz', 'pinho', 'vicente', 'gaspar', 'assunção', 'maia', 'moura', 'valente', 'domingues', 'garcia', 'carneiro', 'loureiro', 'neto', 'amaral', 'branco', 'leal', 'pacheco', 'macedo', 'paiva', 'matias', 'amorim', 'torres', 'adães', 'adorno', 'aguiar',
  'albuquerque', 'alcântara', 'aleluia', 'alencar', 'altamirano', 'alvarenga', 'álvares', 'alvim', 'amigo', 'amor', 'anchieta', 'andrada', 'andré', 'anes', 'antônio', 'anunciação', 'apolinário', 'aragão', 'arruda', 'ascensão', 'assis', 'azeredo', 'bandeira', 'barroso', 'bastos', 'benjamin', 'bermudes', 'bernades', 'bernadino', 'bernardes', 'bernardino', 'bicalho', 'bispo', 'bocaiuva', 'bolsonaro', 'borba', 'borborema', 'botelho', 'braga', 'bragança', 'brasil', 'brasiliense', 'bruno', 'bueno', 'cabral', 'caldas', 'camacho', 'camargo', 'caminha', 'camões',
  'cândido', 'carmo', 'carnaval', 'carvalhal', 'carvalhosa', 'castilho', 'cavalcante', 'chaves', 'coimbra', 'conceição', 'constante', 'cordeiro', 'cotrim', 'coutinho', 'couto', 'curado', 'dambros', 'delfino', 'dias', 'dorneles', 'dourado', 'duarte', 'dutra', 'encarnação', 'evangelista', 'exaltação', 'fagundes', 'falópio', 'falqueto', 'faro', 'figueira', 'fioravante', 'flores', 'fortaleza', 'fracasso', 'frança', 'freire', 'frota', 'furquim', 'furtado', 'gaia', 'galvão', 'gama', 'garrastazu', 'gato', 'generoso', 'gonzaga', 'gouveia', 'guimarães', 'hernandes',
  'holanda', 'homem', 'hora', 'hungria', 'inácio', 'jardim', 'jordão', 'junqueira', 'lacerda', 'lange', 'leitão', 'leme', 'lins', 'lira', 'lisboa', 'luz', 'madureira', 'maduro', 'mairinque', 'malafaia', 'malta', 'marins', 'mascarenhas', 'maurício', 'medeiros', 'médici', 'mendonça', 'menino', 'mesquita', 'messias', 'mioto', 'montenegro', 'moraes', 'morato', 'moro', 'namorado', 'nantes', 'nóbrega', 'noronha', 'oliva', 'padrão', 'paiva', 'paixão', 'papanicolau', 'pascal', 'pascoal', 'patriota', 'peçanha', 'pedrosa', 'pedroso',
  'pensamento', 'penteado', 'peres', 'pessoa', 'pestana', 'pimenta', 'pimentel', 'poeta', 'porto', 'portugal', 'prado', 'prudente', 'peixoto', 'quaresma', 'queiroz', 'ramalhete', 'ramalho', 'ramires', 'rangel', 'rebouças', 'resende', 'roma', 'romão', 'sacramento', 'sales', 'sampaio', 'sampaulo', 'sampedro', 'santacruz', 'santana', 'santander', 'santarrosa', 'santiago', 'saragoça', 'saraiva', 'saramago', 'seixas', 'serra', 'serrano', 'silveira', 'silvério', 'siqueira', 'souza', 'tales', 'teles', 'toledo', 'torquato', 'trindade', 'uchoa', 'uribe',
  'ustra', 'valadares', 'valença', 'valpaços', 'varela', 'vargem', 'vasconcelos', 'veiga', 'veloso', 'veras', 'viana', 'vidal', 'vilhena', 'xavier', 'zampol', 'abaças', 'abacer', 'abade', 'abadeço', 'abadesso', 'abadia', 'abadito', 'abafa', 'abagim', 'abal', 'abalada', 'abalamatos', 'abalde', 'aballi', 'abalo', 'abalroado', 'abana', 'abaraão', 'abarca', 'abascal', 'abavadilha', 'abberasitari', 'abbo', 'abdoulprah', 'abecassis', 'abegão', 'abeijón', 'abel', 'abela', 'abelenda', 'abelha', 'abelhas', 'abelheira', 'abelhi', 'abelho', 'afoito'
].map(s => s.toLowerCase()));

// Nomes que também são palavras comuns (exigem verificação de contexto)
const AMBIGUOUS_PT_NAMES = new Set([
  'pereira', 'afoito', 'rosa', 'serra', 'costa', 'oliveira', 'machado', 'coelho', 'pinto', 'leite', 'maia', 'melo', 'mota', 'neto', 'paiva', 'reis', 'vaz', 'brito', 'cunha', 'guerra', 'matos', 'pinheiro', 'ramalho', 'sequeira', 'abade', 'aguiar', 'aires', 'amado', 'amaral', 'amaro', 'anjos', 'arrais', 'assis', 'azevedo', 'baia', 'bandeira', 'barata', 'barros', 'bastos', 'belo', 'bento', 'bessa', 'bezerra', 'branco', 'brites', 'cabral', 'cacho', 'caldas', 'camacho', 'camara', 'camelo', 'campos', 'canario', 'cano', 'canto', 'capelo', 'carneiro', 'casado', 'casal', 'castro', 'catela', 'chaves', 'cid', 'claro', 'conde', 'coronel', 'crespo', 'cruz', 'curado', 'damas', 'delgado', 'dinis', 'eanes', 'faria', 'farinha', 'faro', 'feio', 'felix', 'ferrao', 'ferro', 'fialho', 'figueira', 'fitas', 'fogaca', 'fontes', 'fraga', 'franca', 'franco', 'freire', 'frois', 'gama', 'garcao', 'garrido', 'girao', 'graca', 'guedes', 'guerreiro', 'homem', 'horta', 'jardim', 'jesus', 'lacerda', 'ladeira', 'lagos', 'lanca', 'leal', 'leao', 'leme', 'lisboa', 'lobo', 'luz', 'magro', 'malta', 'manteigas', 'mariz', 'mata', 'mello', 'mesquita', 'moita', 'moleiro', 'morgado', 'moura', 'mourao', 'nobre', 'novais', 'outeiro', 'padrao', 'paixao', 'palha', 'palma', 'palmeira', 'pardo', 'paredes', 'parreira', 'passos', 'pato', 'paz', 'pena', 'penha', 'pimentel', 'pinho', 'pisco', 'pizarro', 'pocas', 'pontes', 'porto', 'portugal', 'prado', 'prata', 'prego', 'preto', 'pronto', 'rego', 'relvas', 'rico', 'rios', 'rocha', 'rosado', 'rosario', 'sa', 'sales', 'salgado', 'salgueiro', 'salvado', 'santana', 'sardinha', 'seco', 'serpa', 'serrao', 'serrano', 'severo', 'silva', 'silvestre', 'simas', 'sobral', 'sobreira', 'sobrinho', 'solano', 'sousa', 'souto', 'terra', 'toledo', 'torres', 'toscano', 'toste', 'trindade', 'urbano', 'vale', 'valido', 'varela', 'vargas', 'veiga', 'veloso', 'venda', 'ventura', 'viana', 'vidal', 'vieira', 'vilar', 'vinhas'
].map(n => n.toLowerCase()));

// Lista de exceções globais padrão
const DEFAULT_GLOBAL_EXCEPTIONS = [
  'Tribunal da Relação', 'Supremo Tribunal', 'Ministério Público', 'Tribunal Judicial',
  'Comarca de', 'Juízo de', 'Instância Central', 'Instância Local', 'Justiça', 'Direito',
  'Lei', 'Artigo', 'Decreto', 'Portaria', 'Despacho', 'Sentença', 'Acórdão', 'Relatório',
  'Fundamentação', 'Decisão', 'Dispositivo', 'Custas', 'Processo', 'Número', 'Data',
  'Hora', 'Local', 'Sede', 'Empresa', 'Sociedade', 'Limitada', 'Anónima', 'Unipessoal',
  'Herança', 'Jacente', 'Massa', 'Insolvente', 'Falida', 'Estado Português',
  'República Portuguesa', 'Governo', 'Assembleia', 'Câmara Municipal', 'Junta de Freguesia',
  'Região Autónoma', 'Lisboa', 'Porto', 'Coimbra', 'Braga', 'Aveiro', 'Faro', 'Viseu',
  'Évora', 'Guarda', 'Castelo Branco', 'Santarém', 'Setúbal', 'Beja', 'Portalegre',
  'Bragança', 'Vila Real', 'Viana do Castelo', 'Funchal', 'Ponta Delgada',
  'Angra do Heroísmo', 'Horta', 'Banco', 'Sindicato', 'Seguro', 'Refeições', 'Filhos',
  'Social', 'Conta', 'Iban', 'Nif', 'Identificação', 'Fiscal', 'Civil', 'Criminal',
  'Administrativo', 'Fiscal', 'Trabalho', 'Família', 'Menores', 'Comércio', 'Execução',
  'Instrução', 'Criminal', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
  'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo',
  'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira',
  'Norma', 'Comissão', 'Trabalhadores', 'Sucursal', 'Facto', 'Provado',
  'CONSELHEIRO', 'DESEMBARGADOR', 'JUIZ', 'ACORDAM',
  'recorrente', 'recorrida', 'recorrido', 'recorridas', 'recorridos', 'registada', 
  'autor', 'autores', 'autora', 'autoras', 'réu', 'réus', 'ré', 'rés', 
  'requerente', 'requerentes', 'requerida', 'requeridas', 'requerido', 'requeridos',
  'adjunta', 'adjunto', 'desembargadora', 'desembargador', 'conselheira', 'conselheiro',
  'termos em que', 'recurso de apelação', 'termo e duração', 'partes de cima', 'partes de baixo',
  'família de tricotados', 'família dos tricotados', 'rua fernão magalhães',
  'pelo exposto', 'em conformidade', 'nos termos do artigo', 'codigo de processo civil',
  'supremo tribunal de justiça', 'juízo de instrução', 'tribunal da relacao',
  'conforme o disposto', 'nestes termos', 'pede deferimento', 'valor da causa',
  'custas de parte', 'procuradoria', 'taxa de justiça', 'apoio judiciário',
  'acordo e documentos', 'administrativo e fiscal', 'afirmou que foi celebrado',
  'apelações', 'após ser-lhe deferido', 'prorrogação de prazo',
  'assembleia geral', 'assembleia geral da', 'assembleia geral da liga',
  'associação vilanovense', 'pedido de prorrogação', 'para o fazer',
  'Transportes Internacionais Rodoviários', 'concluindo deverem improceder os pedidos do',
  'passou a exercer funções de', 'pagar ao Autor as seguintes quantias',
  'exercer funções de', 'seguintes quantias', 'improceder os pedidos',
  'deverem improceder', 'pagar ao Autor', 'Transportes Internacionais',
  'concluindo deverem', 'passou a exercer', 'funções de', 'seguintes quantias',
  'foi celebrado', 'nos autos', 'em causa', 'no caso', 'de facto', 'de direito',
  'Recursos Humanos', 'Departamento', 'Direção', 'Serviços', 'Fábrica', 'Unidade',
  'Administração', 'Conselho', 'Gerência', 'Secretaria', 'Gabinete', 'Secção',
  'Superior', 'Hierárquico', 'Inferior', 'Colega', 'Trabalhador', 'Funcionário',
  'Colaborador', 'Pessoa', 'Indivíduo', 'Sujeito', 'Cidadão', 'Residente',
  'Morada', 'Rua', 'Avenida', 'Praça', 'Largo', 'Estrada', 'Caminho', 'Beco',
  'Travessa', 'Edifício', 'Prédio', 'Andar', 'Piso', 'Apartamento', 'Fração',
  'Código Postal', 'Localidade', 'Freguesia', 'Concelho', 'Distrito', 'País',
  'Nacionalidade', 'Naturalidade', 'Estado Civil', 'Profissão', 'Habilitações',
  'Contrato', 'Cláusula', 'Anexo', 'Documento', 'Cópia', 'Original', 'Certidão',
  'Registo', 'Notário', 'Conservatória', 'Finanças', 'Segurança Social',
  'Saúde', 'Hospital', 'Centro de Saúde', 'Médico', 'Enfermeiro', 'Paciente',
  'Doença', 'Exame', 'Receita', 'Tratamento', 'Internamento', 'Alta',
  'os', 'as', 'um', 'uma', 'uns', 'umas', 'em', 'no', 'na', 'nos', 'nas',
  'ao', 'à', 'aos', 'às', 'de', 'do', 'da', 'dos', 'das', 'por', 'pelo', 'pela',
  'pelos', 'pelas', 'com', 'sem', 'sob', 'sobre', 'ante', 'após', 'até',
  'desde', 'entre', 'para', 'perante', 'segundo', 'trás', 'atrás', 'dentro',
  'fora', 'cima', 'baixo', 'frente', 'lado', 'meio', 'fim', 'início', 'meio',
  'qual', 'quais', 'quem', 'cujo', 'cuja', 'cujos', 'cujas', 'onde', 'como',
  'quando', 'quanto', 'quanta', 'quantos', 'quantas', 'cada', 'todo', 'toda',
  'todos', 'todas', 'algum', 'alguma', 'alguns', 'algumas', 'nenhum', 'nenhuma',
  'outro', 'outra', 'outros', 'outras', 'mesmo', 'mesma', 'mesmos', 'mesmas',
  'tal', 'tais', 'tanto', 'tanta', 'tantos', 'tantas', 'vários', 'várias',
  'muito', 'muita', 'muitos', 'muitas', 'pouco', 'pouca', 'poucos', 'poucas',
  'mais', 'menos', 'tão', 'quão', 'bastante', 'demais', 'quase', 'apenas',
  'só', 'somente', 'inclusive', 'exclusivamente', 'principalmente', 'especialmente',
  'também', 'ainda', 'já', 'agora', 'logo', 'cedo', 'tarde', 'sempre', 'nunca',
  'jamais', 'talvez', 'provavelmente', 'certamente', 'realmente', 'efetivamente',
  'sim', 'não', 'nem', 'ou', 'mas', 'porém', 'todavia', 'contudo', 'entretanto',
  'logo', 'pois', 'porque', 'visto que', 'já que', 'uma vez que', 'porquanto',
  'conforme', 'segundo', 'consoante', 'embora', 'conquanto', 'ainda que',
  'mesmo que', 'se', 'caso', 'desde que', 'contanto que', 'a menos que',
  'a não ser que', 'para que', 'a fim de que', 'de modo que', 'de sorte que',
  'tanto que', 'tão que', 'como', 'assim como', 'bem como', 'mais que', 'menos que',
  'decorre', 'concordando', 'finalmente', 'começo', 'ficar', 'vender', 'produtos', 'clientes', 'pagou',
  'despesas', 'apresentadas', 'obrigada', 'reembolsá-las', 'comunicou', 'intenção', 'resolução'
];

const ENTITY_BLACKLIST = [
  'BANCO', 'CGD', 'SINDICATO', 'SEGURO', 'REFEICOES', 'FILHOS', 'SOCIAL', 'CONTA', 
  'ESTADO', 'REPUBLICA', 'TRIBUNAL', 'MINISTERIO', 'CONSELHO', 'DIRECAO', 'SERVICO',
  'INSTITUTO', 'AUTORIDADE', 'COMISSAO', 'FUNDACAO', 'ASSOCIACAO', 'FEDERACAO',
  'CONFEDERACAO', 'ORDEM', 'COLEGIO', 'CAMARA', 'JUNTA', 'ASSEMBLEIA',
  'GOVERNO', 'PARLAMENTO', 'PRESIDENCIA', 'SECRETARIA', 'INSPECAO', 'GABINETE',
  'DEPARTAMENTO', 'DIVISAO', 'UNIDADE', 'NUCLEO', 'CENTRO', 'AGENCIA', 'EMPRESA',
  'SOCIEDADE', 'LIMITADA', 'ANONIMA', 'UNIPESSOAL', 'COOPERATIVA', 'MUTUALIDADE',
  'MISERICORDIA', 'SANTA', 'CASA', 'HOSPITAL', 'CLINICA', 'SAUDE',
  'ESCOLA', 'AGRUPAMENTO', 'UNIVERSIDADE', 'FACULDADE', 'POLITECNICO',
  'ACADEMIA', 'CONSERVATORIO', 'BIBLIOTECA', 'MUSEU', 'TEATRO', 'ARQUIVO',
  'DIARIO', 'BOLETIM', 'JORNAL', 'REVISTA', 'TELEVISAO', 'RADIO', 'IMPRENSA',
  'CORREIO', 'TELECOM', 'ENERGIA', 'AGUA', 'SANEAMENTO', 'RESIDUOS', 'TRANSPORTES',
  'METRO', 'COMBOIOS', 'AUTOCARROS', 'AVIAO', 'AEROPORTO', 'PORTO', 'ESTRADA',
  'AUTOESTRADA', 'PONTE', 'TUNEL', 'VIADUTO', 'PARQUE', 'JARDIM', 'FLORESTA',
  'RESERVA', 'MONUMENTO', 'PALACIO', 'CASTELO', 'IGREJA', 'CATEDRAL', 'MOSTEIRO',
  'CONVENTO', 'SANTUARIO', 'ERMIDA', 'CAPELA', 'CEMITERIO', 'ESTADIO', 'PAVILHAO',
  'PISCINA', 'GINASIO', 'CAMPO', 'PISTA', 'CIRCUITO', 'ARENA', 'PRACA', 'AVENIDA',
  'RUA', 'TRAVESSA', 'LARGO', 'BECO', 'CALCADA', 'ESCADA', 'PATIO', 'QUINTA',
  'HERDADE', 'CASAL', 'LUGAR', 'ALDEIA', 'VILA', 'CIDADE', 'CONCELHO', 'DISTRITO',
  'REGIAO', 'PAIS', 'CONTINENTE', 'MUNDO', 'ABRIL', 'DOMINGO', 'NORMA', 'FACTO', 'PROVADO',
  'COMISSAO', 'TRABALHADORES', 'SUCURSAL', 'SOCIEDADE', 'RELATORIO', 'ADVOGADO', 'ADVOGADA',
  'CEDULA', 'PROCESSO', 'PROCEDIMENTO', 'REQUERIMENTO', 'DESPACHO', 'SENTENCA', 'ACORDAO',
  'PAGINA', 'FOLHA', 'DOCUMENTO', 'ANEXO', 'CERTIDAO', 'NOTIFICACAO', 'CITACAO', 'EDITAL',
  'ACORDAM', 'CONSELHEIRO', 'DESEMBARGADOR', 'JUIZ',
  'RECORRENTE', 'RECORRIDA', 'RECORRIDO', 'RECORRIDAS', 'RECORRIDOS', 'REGISTADA', 
  'AUTOR', 'AUTORES', 'AUTORA', 'AUTORAS', 'REU', 'REUS', 'RE', 'RES', 
  'REQUERENTE', 'REQUERENTES', 'REQUERIDA', 'REQUERIDAS', 'REQUERIDO', 'REQUERIDOS',
  'ADJUNTA', 'ADJUNTO', 'DESEMBARGADORA', 'DESEMBARGADOR', 'CONSELHEIRA', 'CONSELHEIRO',
  'TERMOS EM QUE', 'RECURSO DE APELACAO', 'TERMO E DURACAO', 'PARTES DE CIMA', 'PARTES DE BAIXO',
  'FAMILIA DE TRICOTADOS', 'FAMILIA DOS TRICOTADOS', 'RUA FERNAO MAGALHAES',
  'PELO EXPOSTO', 'EM CONFORMIDADE', 'NOS TERMOS DO ARTIGO', 'CODIGO DE PROCESSO CIVIL',
  'NESTES TERMOS', 'PEDE DEFERIMENTO', 'VALOR DA CAUSA', 'TAXA DE JUSTICA',
  'ACORDO', 'DOCUMENTOS', 'ADMINISTRATIVO', 'FISCAL', 'CELEBRADO', 'APELACOES',
  'PRORROGACAO', 'PRAZO', 'ASSEMBLEIA', 'GERAL', 'LIGA', 'ASSOCIACAO', 'VILANOVENSE',
  'PEDIDO', 'DEFERIDO', 'FAZER', 'AUTOS', 'CAUSA', 'CASO', 'FACTO', 'DIREITO', 'JUIZ DE FORA', 'FORA',
  'AUTOR DA HERANÇA', 'AUTOR DO CRIME', 'AUTOR DOS FACTOS', 'AUTORIA', 'PROPRIEDADE INTELECTUAL'
];

const PII_PATTERNS = {
  NIF: /\b[12356789]\d{8}\b/g,
  CC: /\b\d{8}\s*\d\s*[A-Z]{2}\d\b/gi,
  PASSPORT: /\b[A-Z]{1}\d{6}\b/gi,
  PHONE: /\b(?:(?:\+|00)351\s?)?[29]\d{8}\b/g,
  EMAIL: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
  IBAN: /\bPT50\s?\d{21}\b/gi,
  MATRICULA: /\b(?:[A-Z]{2}-\d{2}-\d{2}|\d{2}-\d{2}-[A-Z]{2}|\d{2}-[A-Z]{2}-\d{2}|[A-Z]{2}-\d{2}-[A-Z]{2})\b/g,
  JUIZ: /\bJuiz(?:\(a\))?\s+(?:de\s+Direito\s+)?([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*)/g,
  AUTOR: /\bAutor(?:\(a\))?\s+([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*)/g,
  ADVOGADO: /\b(?:Advogado|Advogada|Mandatário|Mandatária)\s+([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*)/g,
  // More aggressive name patterns for Portuguese - updated to handle unicode boundaries
  NOME_PT: /(?<![a-zA-ZÀ-ÿ0-9_])(?:Sr\.|Sra\.|Dr\.|Dra\.|Eng\.|Prof\.|Juiz|Desembargador|Colega|Autor|Autora|Réu|Ré|Arguido|Arguida|Denunciado|Denunciada|Participante|Mandatário|Advogado|Advogada|Recorrente|Recorrida|Recorrido)(?:,\s*|\s+)([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*)(?![a-zA-ZÀ-ÿ0-9_])/g,
  NOME_CAPS: /(?<![a-zA-ZÀ-ÿ0-9_])([A-ZÀ-ÖØ-Þ]{2,}(?:\s+(?:de|da|do|dos|das|e|DE|DA|DO|DOS|DAS|E)\s+[A-ZÀ-ÖØ-Þ]{2,}|\s+[A-ZÀ-ÖØ-Þ]{2,}){1,8})(?![a-zA-ZÀ-ÿ0-9_])/g,
  // Generic sequence of capitalized words (2 or more) - updated to handle unicode boundaries
  NOME_GENERIC: /(?<![a-zA-ZÀ-ÿ0-9_])([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)){0,8})(?![a-zA-ZÀ-ÿ0-9_])/g,
  // Pattern for names with "e" in the middle (often two people)
  NOME_AND: /(?<![a-zA-ZÀ-ÿ0-9_])([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*\s+e\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.))*)(?![a-zA-ZÀ-ÿ0-9_])/g,
  // Legal context patterns
  NOME_LEGAL: /(?<![a-zA-ZÀ-ÿ0-9_])(?:pelo|pela|por|contra|entre|com|de|do|da|a|ao|à|recorrente|recorrida|recorrido)(?:,\s*|\s+)([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)){0,8})(?![a-zA-ZÀ-ÿ0-9_])/g,
  COLETIVA: /\b(?:Associação|Fundação|Cooperativa|Sociedade|Empresa|Escola|Faculdade|Universidade|Instituto|Centro|Agrupamento|Sindicato|Banco|Seguradora|Companhia|Câmara|Junta|Assembleia|Governo|Estado|República|Ministério|Tribunal|Conselho|Direção|Serviço|Autoridade|Comissão|Unidade|Núcleo|Agência)\s+([A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+|\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){1,8})\b|\b([A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+|\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,8})\s+(?:Lda\.?|Limitada|S\.A\.?|Sociedade\s+Anónima|Unipessoal|S\.?C\.?P\.?|S\.?P\.?Q\.?)\b/g,
};

const NAME_TITLES = [
  'Colega', 'Autor', 'Autora', 'Réu', 'Ré', 'Arguido', 'Arguida', 'Denunciado', 'Denunciada', 'Participante', 'Mandatário', 'Advogado', 'Advogada', 'Recorrente', 'Recorrido',
  'Dr\\.', 'Dra\\.', 'Sr\\.', 'Sra\\.', 'Eng\\.', 'Prof\\.', 'Juiz', 'Desembargador', 'Relator', 
  'Relatora', 'Venerando', 'Tribunal', 'Relação', 'Cfr\\.', 'In', 'Págs\\.', 'Pág\\.', 'Artigo', 'Art\\.', 'N\\.º', 
  'Processo', 'Proc\\.', 'Data', 'Hora', 'Local', 'Sede', 'Empresa', 'Sociedade', 'Trabalhador', 'Trabalhadora',
  'Funcionário', 'Funcionária', 'Agente', 'Cabo', 'Guarda', 'Sargento', 'Tenente', 'Capitão', 'Major', 'Coronel', 'General',
  'Doutor', 'Doutora', 'Senhor', 'Senhora', 'O', 'A', 'Os', 'As', 'Um', 'Uma', 'Página', 'Folha', 'Documento'
];

const NAME_CLEAN_REGEX = new RegExp(`^\\s*(?:${NAME_TITLES.join('|')})\\s+|\\s+(?:${NAME_TITLES.join('|')})\\s*$`, 'gi');
const CONJUNCTION_CLEAN_REGEX = /^\s*(?:e|ou|com|contra)\s+|\s+(?:e|ou|com|contra)\s*$/gi;
const PUNCTUATION_CLEAN_REGEX = /^[.,;:\-\s\(\)\[\]]+|[.,;:\-\s\(\)\[\]]+$/g;

export function isValidNIF(nif: string): boolean {
  const s = nif.replace(/\s/g, '');
  if (!/^[12356789]\d{8}$/.test(s)) return false;
  let checkDigit = 0;
  for (let i = 0; i < 8; i++) {
    checkDigit += parseInt(s[i]) * (9 - i);
  }
  checkDigit = 11 - (checkDigit % 11);
  if (checkDigit >= 10) checkDigit = 0;
  return checkDigit === parseInt(s[8]);
}

export function isValidCC(cc: string): boolean {
  const s = cc.replace(/\s/g, '').toUpperCase();
  if (!/^\d{9}[A-Z]{2}\d$/.test(s)) return false;
  
  const getCharValue = (c: string) => {
    const code = c.charCodeAt(0);
    if (code >= 48 && code <= 57) return code - 48; // 0-9
    return code - 55; // A=10, B=11...
  };

  let sum = 0;
  for (let i = 0; i < s.length; i++) {
    let val = getCharValue(s[i]);
    if (i % 2 === 1) {
      val *= 2;
      if (val > 9) val -= 9;
    }
    sum += val;
  }
  return sum % 10 === 0;
}

export function isValidIBAN(iban: string): boolean {
  const s = iban.replace(/\s/g, '').toUpperCase();
  if (!/^PT50\d{21}$/.test(s)) return false;
  
  // Move PT50 to the end and convert to numbers (P=25, T=29)
  // PT50 -> 25 29 50
  const rearranged = s.substring(4) + "2529" + s.substring(2, 4);
  
  // Modulo 97 using big integers or string manipulation
  let remainder = 0;
  for (let i = 0; i < rearranged.length; i++) {
    remainder = (remainder * 10 + parseInt(rearranged[i])) % 97;
  }
  return remainder === 1;
}

export function cleanName(name: string): string {
  let cleaned = name.trim();
  
  // Repeatedly clean until no more changes (to handle "Colega Dr. António")
  let prev;
  do {
    prev = cleaned;
    cleaned = cleaned.replace(NAME_CLEAN_REGEX, ' ');
    cleaned = cleaned.replace(CONJUNCTION_CLEAN_REGEX, ' ');
    // Handle punctuation at start/end, but be careful with abbreviations like "Dr."
    // We only remove punctuation if it's not part of a known title
    cleaned = cleaned.replace(/^[.,;:\-\s\(\)\[\]]+|[.,;:\-\s\(\)\[\]]+$/g, '');
    cleaned = cleaned.trim();
  } while (cleaned !== prev && cleaned.length > 0);

  return cleaned;
}

export function getNextPseudonym(type: string, existingEntities: PIIEntity[]): string {
  const typeEntities = existingEntities.filter(e => e.type === type);
  const count = new Set(typeEntities.map(e => e.original.toLowerCase())).size + 1;
  
  if (type === 'NOME') {
    // Generate AA, BB, CC...
    const charCode = 64 + count; // 65 is 'A'
    if (count <= 26) {
      const char = String.fromCharCode(charCode);
      return `NOME.${char}${char}`;
    } else {
      // Fallback for more than 26 names: AAA, BBB...
      const repeat = Math.floor((count - 1) / 26) + 2;
      const char = String.fromCharCode(65 + ((count - 1) % 26));
      return `NOME.${char.repeat(repeat)}`;
    }
  }

  const prefixes: Record<string, string> = {
    LOCAL: 'LOCAL',
    PHONE: 'TELEFONE',
    NIF: 'NIF',
    CC: 'CC',
    PASSPORT: 'PASSAPORTE',
    EMAIL: 'EMAIL',
    IBAN: 'IBAN',
    MATRICULA: 'MATRICULA',
    AUTOR: 'AUTOR',
    JUIZ: 'JUIZ',
    ADVOGADO: 'ADVOGADO',
    COLETIVA: 'COLETIVA',
    HEADER: 'CABECALHO',
  };

  const prefix = prefixes[type] || type;
  return `${prefix}.${count}`;
}

// Normalização de texto conforme requisitos
export function normalizeText(text: string): string {
  return text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // Remove acentos
    .replace(/^[.,;:\-\s\(\)\[\]]+|[.,;:\-\s\(\)\[\]]+$/g, '') // Remove pontuação e símbolos nas extremidades
    .replace(/\s+/g, ' ')           // Colapsa espaços múltiplos
    .trim();
}

// Super normalização: remove TUDO exceto letras e números para comparação ultra-robusta
export function superNormalize(text: string): string {
  return text
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

// Distância de Levenshtein para detetar pequenas trocas de letras (typos)
export function levenshteinDistance(a: string, b: string): number {
  const matrix = Array.from({ length: a.length + 1 }, () =>
    Array.from({ length: b.length + 1 }, (_, i) => i)
  );
  for (let i = 1; i <= a.length; i++) matrix[i][0] = i;
  for (let j = 1; j <= b.length; j++) matrix[0][j] = j;
  
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,      // Deletion
        matrix[i][j - 1] + 1,      // Insertion
        matrix[i - 1][j - 1] + cost // Substitution
      );
    }
  }
  return matrix[a.length][b.length];
}

// Função para comparar texto ignorando acentos e case
export function isMatchNormalized(text: string, term: string): boolean {
  const normText = normalizeText(text);
  const normTerm = normalizeText(term);
  return normText === normTerm;
}

export interface Safelist {
  words_ignore: string[];
  phrases_ignore: string[];
}

const NON_NAME_INDICATORS = [
  'transportes', 'internacionais', 'rodoviários', 'concluindo', 'deverem', 'improceder', 
  'pedidos', 'passou', 'exercer', 'funções', 'pagar', 'seguintes', 'quantias',
  'processo', 'tribunal', 'relação', 'supremo', 'justiça', 'artigo', 'número',
  'data', 'hora', 'local', 'sede', 'empresa', 'sociedade', 'geral', 'assembleia',
  'conforme', 'disposto', 'termos', 'pede', 'deferimento', 'valor', 'causa',
  'custas', 'taxa', 'apoio', 'judiciário', 'acordo', 'documentos', 'celebrado',
  'apelações', 'prazo', 'fazer', 'autos', 'facto', 'direito', 'visto', 'identificado',
  'registada', 'notificado', 'citado', 'ofício', 'mandado', 'execução', 'penhora',
  'recorrente', 'recorrido', 'autor', 'réu', 'requerente', 'requerido', 'apelante', 'apelado',
  'oposição', 'embargos', 'executado', 'exequente', 'testemunha', 'perito', 'escrivão',
  'conclui', 'decide', 'julga', 'improcedente', 'procedente', 'parcialmente', 'absolve', 
  'condena', 'recurso', 'apelação', 'revista', 'agravo', 'reclamação', 'contestação', 
  'réplica', 'tréplica', 'alegações', 'conclusões', 'fundamentos', 'factos', 'provados', 
  'aplicável', 'decisão', 'final', 'vistos', 'examinados', 'relatados', 'discutidos',
  'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
  'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado', 'domingo',
  'agravamento', 'patologia', 'anca', 'direita', 'consequência', 'direta', 'lesão', 'bacia', 'provocada', 'pelo', 'pela', 'at', 'de', 'do', 'da', 'dos', 'das',
  'clínico', 'médico', 'hospital', 'saúde', 'exame', 'diagnóstico', 'tratamento', 'sintomas', 'doença', 'lesões', 'ferimentos', 'acidente', 'trabalho',
  'comunicou', 'caducidade', 'detrimento', 'indemnização', 'vestuário', 'perspetivando', 'cofemel', 'sociedade', 'empresa', 'marca', 'modelo', 'desenho',
  'propriedade', 'intelectual', 'industrial', 'direitos', 'autor', 'obra', 'original', 'criação', 'proteção', 'âmbito', 'aplicação', 'diretiva', 'regulamento',
  'acórdão', 'sentença', 'decisão', 'fundamentação', 'factos', 'provados', 'não', 'provados', 'direito', 'aplicável', 'dispositivo', 'custas', 'valor', 'causa',
  'recurso', 'apelação', 'revista', 'agravo', 'reclamação', 'oposição', 'embargos', 'execução', 'penhora', 'citação', 'notificação', 'prazo', 'requerimento',
  'automóveis', 'automóvel', 'veículo', 'viatura', 'estrada', 'trânsito', 'condução', 'carta', 'seguro', 'apólice', 'sinistro',
  'pagar', 'receber', 'entregar', 'liquidar', 'vencer', 'vencimento', 'juros', 'mora', 'legal', 'capital', 'quantia', 'montante',
  'reintegrar', 'reintegração', 'restituir', 'restituição', 'indemnizar', 'compensar', 'reparar', 'danos', 'prejuízos',
  'lda', 'ld', 'sa', 'unipessoal', 'limitada', 'sociedade', 'gerência', 'administração', 'direção', 'conselho', 'fiscal',
  'comunicar', 'notificar', 'citar', 'informar', 'esclarecer', 'declarar', 'atestar', 'certificar', 'confirmar'
];

function calculateNameScore(
  text: string, 
  type: string, 
  contextBefore: string, 
  contextAfter: string, 
  repetitionCount: number
): number {
  let score = 0;
  const norm = normalizeText(text);
  const allWords = norm.split(/\s+/).filter(w => w.length > 0);
  const trimmed = text.trim();
  
  // 0. Check for verbs (common endings in Portuguese)
  // Added infinitive endings: ar, er, ir
  const verbEndings = ['ando', 'endo', 'indo', 'aram', 'eram', 'iram', 'ou', 'amos', 'emos', 'imos', 'ará', 'erá', 'irá', 'ar', 'er', 'ir'];
  const hasVerbEnding = allWords.some(w => {
    const low = w.toLowerCase();
    // Only check if it's not a common name (e.g., "Aguiar" ends in "ar" but is a name)
    return low.length > 4 && verbEndings.some(e => low.endsWith(e)) && !COMMON_PT_FIRST_NAMES.has(low) && !COMMON_PT_SURNAMES.has(low);
  });
  if (hasVerbEnding) score -= 8; // Increased penalty

  // 0.1 Check for common abbreviations that are not names
  const commonAbbr = ['ld', 'lda', 'sa', 'nº', 'n.º', 'art', 'art.º', 'pág', 'fls', 'doc'];
  if (allWords.some(w => commonAbbr.includes(w.toLowerCase()))) score -= 10;

  // 1. Positive triggers in contextBefore
  const positiveTriggers = [
    'nome', 'requerente', 'réu', 'ré', 'autor', 'autora', 'mandatário', 'mandatária', 
    'dr.', 'dra.', 'sr.', 'sra.', 'exmo.', 'exma.', 'trabalhador', 'trabalhadora', 
    'trabalhadores', 'trabalhados', 'participante', 'funcionário', 'funcionária',
    'identificado', 'identificada', 'apelante', 'apelado', 'condutor', 'condutora',
    'proprietário', 'proprietária', 'vítima', 'ofendido', 'ofendida', 'assistente',
    'testemunha', 'depoente', 'declarante', 'perito', 'perita'
  ];
  const contextLower = contextBefore.toLowerCase();
  const contextAfterLower = contextAfter.toLowerCase();
  
  if (positiveTriggers.some(t => contextLower.includes(t) || contextAfterLower.includes(t))) {
    score += 5; // Increased boost
  }

  // 2. Dictionary matches
  let dictionaryMatches = 0;
  allWords.forEach(word => {
    const w = word.toLowerCase();
    if (COMMON_PT_FIRST_NAMES.has(w) || COMMON_PT_SURNAMES.has(w)) {
      dictionaryMatches++;
    }
  });
  
  const matchRatio = dictionaryMatches / allWords.length;
  
  if (dictionaryMatches >= 2) {
    score += 4;
    if (matchRatio > 0.6) score += 2;
  } else if (dictionaryMatches === 1) {
    score += 1;
  } else {
    score -= 6; // Increased penalty for zero dictionary matches
  }

  // 3. Repetition
  if (repetitionCount >= 3) score += 2;
  else if (repetitionCount === 2) score += 1;

  // 4. Negative indicators (Institutional/Medical/Legal terms)
  const foundNegativeIndicators = allWords.filter(w => NON_NAME_INDICATORS.includes(w.toLowerCase()));
  if (foundNegativeIndicators.length > 0) {
    score -= (6 * foundNegativeIndicators.length); // Increased penalty per indicator
  }

  // 5. Address/Postal code context
  const addressTriggers = ['rua', 'avenida', 'praça', 'largo', 'estrada', 'caminho', 'beco', 'travessa', 'n.º', 'nº', 'código postal', 'cp'];
  if (addressTriggers.some(t => contextLower.includes(t) || contextAfterLower.includes(t))) {
    score -= 4;
  }

  // 6. All-caps penalty
  const isAllCaps = trimmed === trimmed.toUpperCase() && trimmed.length > 5;
  if (isAllCaps) {
    if (dictionaryMatches === 0) score -= 10; // Even heavier penalty
    else if (matchRatio < 0.5) score -= 5;
  }

  // 7. Length penalty
  if (allWords.length === 1 && dictionaryMatches === 0) score -= 8;
  if (allWords.length > 5) score -= 5;
  if (allWords.length > 8) score -= 10;

  // 8. Start/End with prepositions/conjunctions/verbs
  const connectors = ['de', 'do', 'da', 'dos', 'das', 'e', 'com', 'contra', 'pelo', 'pela', 'ao', 'à', 'em', 'por', 'para', 'o', 'a', 'os', 'as'];
  if (connectors.includes(allWords[0]?.toLowerCase())) score -= 5;
  if (connectors.includes(allWords[allWords.length - 1]?.toLowerCase())) score -= 5;

  // 9. Conector bonus (e.g., "de", "da" in the middle of a name is a strong indicator)
  const hasInternalConnector = allWords.slice(1, -1).some(w => connectors.includes(w.toLowerCase()));
  if (hasInternalConnector && allWords.length >= 3) {
    score += 4;
  }

  // 10. Lowercase start (if not a connector)
  if (trimmed[0] === trimmed[0].toLowerCase() && !connectors.includes(allWords[0]?.toLowerCase())) {
    score -= 6;
  }

  // 10. Ambiguous name check (e.g., "Afoito" or "Pereira" alone)
  if (allWords.length === 1 && AMBIGUOUS_PT_NAMES.has(allWords[0].toLowerCase())) {
    // Only penalize if it doesn't have a title context
    const contextLower = contextBefore.toLowerCase();
    const titles = ['dr.', 'dra.', 'sr.', 'sra.', 'exmo.', 'exma.', 'juiz', 'autor', 'réu', 'arguido', 'arguida'];
    if (!titles.some(t => contextLower.includes(t))) {
      score -= 5;
    }
  }

  return score;
}

export function scanText(
  text: string, 
  fileId: string, 
  existingEntities: PIIEntity[] = [], 
  isRelated: boolean = true, 
  globalKnowledge: Record<string, string> = {},
  safelist: Safelist = { words_ignore: [], phrases_ignore: [] },
  positions: { start: number, end: number, page: number, y: number }[] = []
): PIIEntity[] {
  const entities: PIIEntity[] = [];
  const foundMatches: { text: string, type: string, start: number, end: number, reason?: string, page?: number, y?: number }[] = [];

  // Helper to get position for a text range
  const getPosition = (start: number, end: number) => {
    // Find the paragraph that contains the start of the match
    const pos = positions.find(p => start >= p.start && start <= p.end);
    return pos || null;
  };

  // Pré-processamento: Identificar áreas protegidas pela Safelist
  const protectedRanges: { start: number, end: number, term: string }[] = [];
  
  // Normalizar o texto mantendo o mapeamento de índices (aproximado, lidando com acentos)
  // Para simplificar e manter precisão, vamos usar regex que ignoram acentos
  const getRegexForTerm = (term: string) => {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // Substituir letras com acentos por classes de caracteres
    const withAccents = escaped
      .replace(/[aáàâã]/gi, '[aáàâã]')
      .replace(/[eéèê]/gi, '[eéèê]')
      .replace(/[iíìî]/gi, '[iíìî]')
      .replace(/[oóòôõ]/gi, '[oóòôõ]')
      .replace(/[uúùû]/gi, '[uúùû]')
      .replace(/[cç]/gi, '[cç]');
    
    // Permitir múltiplos espaços entre palavras
    const withSpaces = withAccents.replace(/\s+/g, '\\s+');
    // Usar lookarounds em vez de \b para suporte a Unicode (acentos)
    // (?<![a-zA-ZÀ-ÿ]) garante que não é precedido por uma letra
    // (?![a-zA-ZÀ-ÿ]) garante que não é seguido por uma letra
    return new RegExp(`(?<![a-zA-ZÀ-ÿ])${withSpaces}(?![a-zA-ZÀ-ÿ])`, 'gi');
  };

  // PASSO A: Expressões a ignorar
  safelist.phrases_ignore.forEach(phrase => {
    const regex = getRegexForTerm(phrase);
    let match;
    while ((match = regex.exec(text)) !== null) {
      protectedRanges.push({ start: match.index, end: match.index + match[0].length, term: phrase });
    }
  });

  // PRESERVAR TRATAMENTOS E DIVISÕES MANUAIS ANTERIORES:
  // Se uma entidade já foi marcada como tratada (validada/anonimizada), ignorada ou pertence a um grupo manual, 
  // nós devemos encontrar todas as suas ocorrências neste documento e protegê-las.
  // Isso impede que padrões automáticos gananciosos (como NOME_AND) englobem ou desfaçam as divisões manuais do utilizador.
  existingEntities.forEach(existing => {
    if (existing.treated || existing.ignored || existing.groupId?.startsWith('manual-group-')) {
      const regex = getRegexForTerm(existing.original);
      let match;
      while ((match = regex.exec(text)) !== null) {
        const start = match.index;
        const end = start + match[0].length;
        const pos = getPosition(start, end);

        // Adicionar de volta aos matches encontrados para que o rescan continue a reportá-lo
        foundMatches.push({
          text: match[0],
          type: existing.type,
          start: start,
          end: end,
          page: pos?.page,
          y: pos?.y
        });

        // Proteger esta região para que nenhum outro padrão (ex: NOME_AND) se sobreponha a ela
        protectedRanges.push({
          start: start,
          end: end,
          term: existing.original
        });
      }
    }
  });

  // PASSO B: Palavras a ignorar (serão usadas no filtro final dos matches)
  const normalizedWordsIgnore = new Set(safelist.words_ignore.map(w => normalizeText(w)));

  // PASSO C: Conhecimento Global normalizado para busca rápida
  const normalizedKnowledge = new Map<string, string>();
  const superNormalizedExceptions = new Set<string>();
  
  // Adicionar palavras e frases da safelist às exceções super normalizadas
  safelist.words_ignore.forEach(w => superNormalizedExceptions.add(superNormalize(w)));
  safelist.phrases_ignore.forEach(p => superNormalizedExceptions.add(superNormalize(p)));

  Object.entries(globalKnowledge).forEach(([k, t]) => {
    const norm = normalizeText(k);
    normalizedKnowledge.set(norm, t);
    if (t === 'EXCECAO') {
      superNormalizedExceptions.add(superNormalize(k));
    }
  });

  // Helper para verificar se um texto é uma exceção de forma robusta
  const isException = (matchText: string): boolean => {
    const norm = normalizeText(matchText);
    if (normalizedWordsIgnore.has(norm)) return true;
    if (normalizedKnowledge.get(norm) === 'EXCECAO') return true;

    const snorm = superNormalize(matchText);
    if (superNormalizedExceptions.has(snorm)) return true;

    // Fuzzy matching para pequenas trocas de letras (typos)
    // Apenas para termos com comprimento razoável para evitar falsos positivos
    if (snorm.length >= 3) {
      for (const ex of superNormalizedExceptions) {
        // Se a diferença de tamanho for pequena e a distância for 1, consideramos match
        if (Math.abs(ex.length - snorm.length) <= 1) {
          if (levenshteinDistance(ex, snorm) <= 1) return true;
        }
      }
    }

    // Caso especial: Se for um nome composto (ex: "I. Pelos"), verificar se a parte principal é exceção
    // REMOVIDO: Era demasiado agressivo e ignorava nomes válidos que continham palavras comuns (ex: "António Lisboa")
    /*
    if (snorm.length > 3) {
      const parts = matchText.split(/[\s,.]+/).filter(p => p.length > 2);
      
      // Lista de títulos normalizados para ignorar na verificação de exceção
      const normalizedTitles = NAME_TITLES.map(t => superNormalize(t.replace('\\.', '.')));
      
      for (const part of parts) {
        const sp = superNormalize(part);
        // Se a parte for um título (como "Autora"), não a usamos para invalidar o match
        if (normalizedTitles.includes(sp)) continue;
        
        if (superNormalizedExceptions.has(sp)) return true;
      }
    }
    */

    return false;
  };

  // PASSO D: Identificar nomes conhecidos no texto antes de outros padrões
  // Isto garante que nomes completos sejam capturados mesmo que o NLP falhe
  Object.entries(globalKnowledge).forEach(([name, type]) => {
    if (type !== 'EXCECAO') {
      const regex = getRegexForTerm(name);
      let match;
      while ((match = regex.exec(text)) !== null) {
        const start = match.index;
        const end = match.index + match[0].length;
        const pos = getPosition(start, end);

        // CRITICAL: HEADER type is only valid at the top of the page (top 20%)
        if (type === 'HEADER' && pos && pos.y > 0.20) {
          continue;
        }

        foundMatches.push({
          text: match[0],
          type: type,
          start: start,
          end: end,
          reason: 'global-knowledge',
          page: pos?.page,
          y: pos?.y
        });
      }
    }
  });

  // 1. Regex Patterns
    Object.entries(PII_PATTERNS).forEach(([type, pattern]) => {
      let match;
      const regex = new RegExp(pattern.source, pattern.flags);
      while ((match = regex.exec(text)) !== null) {
        const matchText = match[1] || match[0];
        
        // Checksum validation
        if (type === 'NIF' && !isValidNIF(matchText)) continue;
        if (type === 'CC' && !isValidCC(matchText)) continue;
        if (type === 'IBAN' && !isValidIBAN(matchText)) continue;

        const start = match.index + (match[0].indexOf(matchText));
        const end = start + matchText.length;
        const pos = getPosition(start, end);

        // CRITICAL: HEADER type is only valid at the top of the page
        if (type === 'HEADER' && pos && pos.y > 0.20) {
          continue;
        }

        // Verificar se está em área protegida (PASSO A) - Melhorado para detetar sobreposições
        if (protectedRanges.some(r => start < r.end && end > r.start)) continue;

        // Verificação robusta de exceções (PASSO B, C e Fuzzy)
        if (isException(matchText)) continue;

        foundMatches.push({
          text: matchText,
          type: type.startsWith('NOME') ? 'NOME' : type,
          start: start,
          end: end,
          page: pos?.page,
          y: pos?.y
        });
      }
    });

  // 2. Portuguese Legal Patterns (Parties) - Updated to require capitalization for names and handle internal spaces
  const legalPatterns = [
    /(?:Recorrente|Recorrida|Recorrido|Requerente|Requerido|Réu|Ré|Arguido|Arguida|Denunciado|Denunciada|Participante|Assistente|Beneficiário|Executado|Exequente|Oponente|Reclamante|Reclamado|Interveniente|Contrainteressado|Apelante|Apelado|Agravante|Agravado|Embargante|Embargado|Demandante|Demandado|Advogado|Advogada|Mandatário|Mandatária|Autor|Autora)(?::|,\s*|\s+)\s*([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)){0,8})/g,
    /(?:Nome|Apelido|Filiação|Naturalidade|Residência|Sede)(?::|,\s*|\s+)\s*([A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)(?:\s+(?:de|da|do|dos|das|e)\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)|\s+[A-ZÀ-ÖØ-Þ](?:\s*[a-zà-öø-ÿ]+|\.)){0,8})/g,
  ];

  legalPatterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      if (match[1]) {
        const matchText = match[1].trim();
        const index = match.index + match[0].indexOf(matchText);
        const end = index + matchText.length;
        const pos = getPosition(index, end);

        // Verificar Safelist - Melhorado para detetar sobreposições
        if (protectedRanges.some(r => index < r.end && end > r.start)) continue;
        
        // Verificação robusta de exceções
        if (isException(matchText)) continue;
        
        // Ensure the matched text looks like a name/entity (at least one capital letter)
        // This avoids catching common lowercase phrases like "acordo e documentos"
        if (!/[A-ZÀ-Ÿ]/.test(matchText)) continue;
        
        let type = 'NOME';
        const prefix = match[0].split(':')[0].toLowerCase();
        const autorPrefixes = [
          'recorrente', 'recorrido', 'requerente', 'requerido', 'réu', 'participante', 
          'denunciado', 'arguido', 'assistente', 'beneficiário', 'executado', 'exequente', 
          'oponente', 'reclamante', 'reclamado', 'interveniente', 'contrainteressado', 
          'apelante', 'apelado', 'agravante', 'agravado', 'embargante', 'embargado', 
          'demandante', 'demandado'
        ];
        
        if (autorPrefixes.some(p => prefix.includes(p))) {
          type = 'AUTOR';
        } else if (prefix.includes('juiz') || prefix.includes('desembargador')) {
          type = 'JUIZ';
        } else if (prefix.includes('advogado') || prefix.includes('advogada') || prefix.includes('mandatário') || prefix.includes('mandatária')) {
          type = 'ADVOGADO';
        }

        foundMatches.push({
          text: matchText,
          type: type,
          start: index,
          end: end,
          page: pos?.page,
          y: pos?.y
        });
      }
    }
  });

  // 4. Merge and Deduplicate
  const sortedMatches = foundMatches.sort((a, b) => {
    if (a.start !== b.start) return a.start - b.start;
    return b.text.length - a.text.length;
  });
  
  const mergedMatches: typeof foundMatches = [];
  
  sortedMatches.forEach(match => {
    if (mergedMatches.length === 0) {
      mergedMatches.push(match);
      return;
    }
    
    const last = mergedMatches[mergedMatches.length - 1];
    // Se o match atual está contido no anterior, ignorar
    if (match.start >= last.start && match.end <= last.end) return;
    
    // Se há sobreposição
    if (match.start < last.end) {
      // Se o novo é significativamente mais longo ou começa na mesma posição e é mais longo
      if (match.text.length > last.text.length) {
        mergedMatches[mergedMatches.length - 1] = match;
      }
      return;
    }
    
    mergedMatches.push(match);
  });

  // Count occurrences for repetition scoring
  const occurrenceCounts = new Map<string, number>();
  mergedMatches.forEach(m => {
    const norm = normalizeText(m.text);
    occurrenceCounts.set(norm, (occurrenceCounts.get(norm) || 0) + 1);
  });

  const addEntity = (original: string, type: string, start: number, end: number, page?: number, y?: number) => {
    let trimmed = original.trim();
    
    // Clean names specifically
    if (type === 'NOME' || type === 'AUTOR' || type === 'JUIZ' || type === 'ADVOGADO') {
      trimmed = cleanName(trimmed);
    } else {
      trimmed = trimmed.replace(/[.,;]+$/, '');
    }

    if (trimmed.length < 2) return;

    const lower = trimmed.toLowerCase();
    const norm = normalizeText(trimmed);
    
    // PASSO B (Reforço): Verificar Safelist novamente após limpeza
    if (normalizedWordsIgnore.has(norm)) return;

    // PASSO C: Listas de anonimização forçada (Exceções Globais)
    // Se estiver nas exceções globais (que agora podem vir do conhecimento global), ignorar.
    const knowledgeType = normalizedKnowledge.get(norm);
    if (knowledgeType === 'EXCECAO') return;
    
    // Check if exists in existing entities (for grouping and property preservation)
    // Prioritize finding by text to respect user decisions even if type would be different
    const existing = existingEntities.find(e => normalizeText(e.original) === norm);

    // Determine the type: 
    // 1. If existing is treated/ignored, use its type
    // 2. Otherwise use knowledgeType if available
    // 3. Otherwise use the detected type
    let identifiedType = (existing && (existing.treated || existing.ignored)) 
      ? existing.type 
      : (knowledgeType || type);

    // Check if we already added this entity in the CURRENT scan of this file
    // This prevents duplicate entities for multiple occurrences in the same file
    const alreadyAdded = entities.find(e => normalizeText(e.original) === norm && e.type === identifiedType);
    if (alreadyAdded) {
      // Just ensure the fileId is included (should already be there)
      if (!alreadyAdded.fileIds?.includes(fileId)) {
        alreadyAdded.fileIds = [...(alreadyAdded.fileIds || []), fileId];
      }
      return;
    }

    if (DEFAULT_GLOBAL_EXCEPTIONS.some(ex => normalizeText(ex) === norm)) return;
    if (ENTITY_BLACKLIST.includes(norm.toUpperCase())) return;

    // Advanced Judge Identification based on globalKnowledge
    // ONLY if not already decided by user (treated) or explicit knowledge
    const before = text.substring(Math.max(0, start - 300), start);
    const after = text.substring(end, Math.min(text.length, end + 300));
    
    // Capture more words for better context in modals
    const wordsBefore = before.trim().split(/\s+/).slice(-15).join(' ');
    const wordsAfter = after.trim().split(/\s+/).slice(0, 15).join(' ');
    
    // Context snippet for quick lists (2 words before + entity + 2 words after)
    const contextWordsBefore = before.trim().split(/\s+/).slice(-2).join(' ');
    const contextWordsAfter = after.trim().split(/\s+/).slice(0, 2).join(' ');
    const contextSnippet = `${contextWordsBefore} ${trimmed} ${contextWordsAfter}`.trim();

    // SCORING PIPELINE
    let score = 10; // Default high for non-name types (NIF, EMAIL, etc)
    const isNameType = identifiedType === 'NOME' || identifiedType === 'AUTOR' || identifiedType === 'JUIZ' || identifiedType === 'ADVOGADO';
    
    if (isNameType) {
      const repetitionCount = occurrenceCounts.get(norm) || 1;
      score = calculateNameScore(trimmed, identifiedType, wordsBefore, wordsAfter, repetitionCount);

      // Skip very low confidence names to avoid cluttering the list with false positives
      if (score < -2) return;
    }

    // Advanced Judge Identification based on globalKnowledge
    // ONLY if not already decided by user (treated) or explicit knowledge
    const nameWords = norm.split(/\s+/).filter(w => w.length > 2);

    if (!knowledgeType && (!existing || (!existing.treated && !existing.ignored)) && 
        identifiedType !== 'JUIZ' && identifiedType !== 'AUTOR' && nameWords.length >= 2) {
      // Calcular scores de correspondência para Juízes e Autores
      let bestJudgeScore = 0;
      let bestAuthorScore = 0;
      
      const judges: string[] = [];
      const authors: string[] = [];
      normalizedKnowledge.forEach((t, k) => {
        if (t === 'JUIZ') judges.push(k);
        if (t === 'AUTOR') authors.push(k);
      });

      judges.forEach(judgeName => {
        const judgeWords = judgeName.split(/\s+/).filter(w => w.length > 2);
        if (judgeWords.length < 2) return;
        const common = nameWords.filter(w => judgeWords.includes(w));
        if (common.length >= 2) {
          const score = common.length / Math.max(nameWords.length, judgeWords.length);
          if (score > bestJudgeScore) bestJudgeScore = score;
        }
      });

      authors.forEach(authorName => {
        const authorWords = authorName.split(/\s+/).filter(w => w.length > 2);
        if (authorWords.length < 2) return;
        const common = nameWords.filter(w => authorWords.includes(w));
        if (common.length >= 2) {
          const score = common.length / Math.max(nameWords.length, authorWords.length);
          if (score > bestAuthorScore) bestAuthorScore = score;
        }
      });

      // Atribuir o tipo com melhor correspondência (mínimo de 2 palavras e score > 0.7)
      // O utilizador pediu rigor: se tem duas palavras iguais, deve sugerir (score boost)
      if (bestJudgeScore > 0.4 || bestAuthorScore > 0.4) {
        if (bestJudgeScore >= bestAuthorScore) {
          identifiedType = 'JUIZ';
        } else {
          identifiedType = 'AUTOR';
        }
        score += 10; // Boost score significantly if matched with global knowledge
      }
    }

    if (existing && (existing.type === identifiedType || existing.treated || existing.ignored)) {
      // Return a copy with updated fileIds if needed
      const updatedFileIds = existing.fileIds?.includes(fileId) 
        ? existing.fileIds 
        : [...(existing.fileIds || []), fileId];
        
      entities.push({
        ...existing,
        type: identifiedType, // Update type if it was refined but not treated
        pseudonym: existing.type !== identifiedType ? getNextPseudonym(identifiedType, [...existingEntities, ...entities]) : existing.pseudonym,
        fileIds: updatedFileIds,
        score: score,
        page: page || existing.page,
        y: y || existing.y
      });
      return;
    }

    const id = generateId();
    const pseudonym = getNextPseudonym(identifiedType, [...existingEntities, ...entities]);
    
    // DECISION: Enable automatically only if score is high enough
    const isAutoEnabled = score >= 4;

    entities.push({
      id,
      original: trimmed,
      type: identifiedType,
      pseudonym,
      enabled: isAutoEnabled,
      fileIds: [fileId],
      context: wordsBefore.split(/\s+/).pop() || '',
      contextBefore: wordsBefore,
      contextAfter: wordsAfter,
      contextSnippet,
      score: score,
      page,
      y
    });
  };

  mergedMatches.forEach(m => {
    addEntity(m.text, m.type, m.start, m.end, m.page, m.y);
  });

  return entities;
}

export function splitEntity(entity: PIIEntity, entities: PIIEntity[]): PIIEntity[] {
  const words = entity.original.split(/\s+/).filter(w => w.length >= 2);
  if (words.length < 2) return [entity];

  return words.map(word => {
    const pseudonym = getNextPseudonym(entity.type, entities);
    return {
      ...entity,
      id: generateId(),
      original: word,
      pseudonym,
      groupId: undefined
    };
  });
}

export function groupSimilarEntities(entities: PIIEntity[], isRelated: boolean = true): PIIEntity[] {
  // 1. Identify manual groups and preserve them
  const newEntities = entities.map(e => ({ 
    ...e, 
    groupId: e.groupId?.startsWith('manual-group-') ? e.groupId : undefined 
  }));
  
  // Pre-calculate counts for "at least 2" rule using superNormalize
  const counts: Record<string, number> = {};
  newEntities.forEach(entity => {
    const fileId = entity.fileIds?.[0] || 'unknown';
    const sNorm = superNormalize(entity.original);
    const key = isRelated 
      ? `${entity.type}:${sNorm}`
      : `${fileId}:${entity.type}:${sNorm}`;
    counts[key] = (counts[key] || 0) + 1;
  });

  // 2. Group by exact match (normalized) for all types
  const groups: Record<string, { id: string, pseudonym: string, treated: boolean, type: string }> = {};
  
  // Sort: Manual groups first, then treated, then others
  const sortedEntities = [...newEntities].sort((a, b) => {
    const aManual = a.groupId?.startsWith('manual-group-') ? 1 : 0;
    const bManual = b.groupId?.startsWith('manual-group-') ? 1 : 0;
    if (aManual !== bManual) return bManual - aManual;
    
    const aTreated = a.treated ? 1 : 0;
    const bTreated = b.treated ? 1 : 0;
    return bTreated - aTreated;
  });
  
  sortedEntities.forEach(entity => {
    const fileId = entity.fileIds?.[0] || 'unknown';
    const sNorm = superNormalize(entity.original);
    const key = isRelated 
      ? `${entity.type}:${sNorm}`
      : `${fileId}:${entity.type}:${sNorm}`;

    // Only create a group if there's more than one occurrence OR it's already a manual group
    const hasMultiple = counts[key] > 1;
    const isManual = entity.groupId?.startsWith('manual-group-');

    if (hasMultiple || isManual) {
      if (!groups[key]) {
        groups[key] = { 
          id: entity.groupId || `group-${entity.id}`, 
          pseudonym: entity.pseudonym,
          treated: entity.treated || false,
          type: entity.type
        };
      } else {
        if (entity.groupId?.startsWith('manual-group-') && !groups[key].id.startsWith('manual-group-')) {
          groups[key].id = entity.groupId;
          groups[key].pseudonym = entity.pseudonym;
          groups[key].treated = true;
        } else if (entity.treated && !groups[key].treated) {
          groups[key].pseudonym = entity.pseudonym;
          groups[key].treated = true;
          groups[key].type = entity.type;
        }
      }
      
      const originalEntity = newEntities.find(e => e.id === entity.id);
      if (originalEntity) {
        originalEntity.groupId = groups[key].id;
        originalEntity.pseudonym = groups[key].pseudonym;
      }
    } else {
      // Single occurrence and not manual: ensure no groupId
      const originalEntity = newEntities.find(e => e.id === entity.id);
      if (originalEntity) {
        originalEntity.groupId = undefined;
      }
    }
  });

  // 3. Special handling for NAMES (partial matches and shared words)
  const nameEntities = newEntities.filter(e => e.type === 'NOME' || e.type === 'AUTOR' || e.type === 'JUIZ');
  const getWords = (text: string) => 
    text.toLowerCase().split(/\s+/).filter(w => w.length > 2 && !NAME_TITLES.some(t => new RegExp(`^${t.replace('.', '\\.')}$`, 'i').test(w)));

  const VERY_COMMON_NAMES = new Set(['maria', 'jose', 'manuel', 'antonio', 'joao', 'francisco', 'carlos', 'paulo', 'pedro', 'luis', 'ana', 'isabel', 'teresa', 'margarida', 'silva', 'santos', 'ferreira', 'pereira', 'oliveira', 'costa', 'rodrigues', 'martins', 'jesus']);

  // Optimization: Use an adjacency list and find connected components
  const adj = new Map<number, number[]>();
  
  for (let i = 0; i < nameEntities.length; i++) {
    for (let j = i + 1; j < nameEntities.length; j++) {
      const e1 = nameEntities[i];
      const e2 = nameEntities[j];
      
      if (!isRelated && e1.fileIds?.[0] !== e2.fileIds?.[0]) continue;

      let isMatch = false;
      
      // Use superNormalize to handle "S ara" vs "Sara"
      if (superNormalize(e1.original) === superNormalize(e2.original)) {
        isMatch = true;
      }

      if (!isMatch) {
        const w1 = getWords(e1.original);
        const w2 = getWords(e2.original);
        if (w1.length > 0 && w2.length > 0) {
          const common = w1.filter(w => w2.includes(w));
          const commonNonGeneric = common.filter(w => !VERY_COMMON_NAMES.has(w));
          const similarity = common.length / Math.max(w1.length, w2.length);
          
          if (similarity >= 0.8) {
            isMatch = true;
          } else if (w1.every(w => w2.includes(w)) || w2.every(w => w1.includes(w))) {
            if (common.length >= 2 && commonNonGeneric.length >= 1) {
              if (Math.abs(w1.length - w2.length) <= 2) {
                isMatch = true;
              }
            }
          } else if (common.length >= 3 && commonNonGeneric.length >= 2) {
            isMatch = true;
          }
        }
      }

      const hasSeparator1 = e1.original.toLowerCase().includes(' e ');
      const hasSeparator2 = e2.original.toLowerCase().includes(' e ');
      if (hasSeparator1 !== hasSeparator2) isMatch = false;

      if (isMatch) {
        if (!adj.has(i)) adj.set(i, []);
        if (!adj.has(j)) adj.set(j, []);
        adj.get(i)!.push(j);
        adj.get(j)!.push(i);
      }
    }
  }

  const visited = new Set<number>();
  for (let i = 0; i < nameEntities.length; i++) {
    if (visited.has(i)) continue;
    
    const component: number[] = [];
    const queue = [i];
    visited.add(i);
    
    while (queue.length > 0) {
      const curr = queue.shift()!;
      component.push(curr);
      const neighbors = adj.get(curr) || [];
      for (const neighbor of neighbors) {
        if (!visited.has(neighbor)) {
          visited.add(neighbor);
          queue.push(neighbor);
        }
      }
    }

    if (component.length > 1) {
      let targetG: string | undefined;
      let targetP: string | undefined;
      let isTreated = false;

      for (const idx of component) {
        const e = nameEntities[idx];
        if (e.groupId?.startsWith('manual-group-')) {
          targetG = e.groupId;
          targetP = e.pseudonym;
          isTreated = true;
          break;
        }
        if (e.treated && !isTreated) {
          targetG = e.groupId;
          targetP = e.pseudonym;
          isTreated = true;
        }
      }

      if (!targetG) {
        const longest = component.reduce((prev, curr) => 
          nameEntities[curr].original.length > nameEntities[prev].original.length ? curr : prev, component[0]);
        targetG = nameEntities[longest].groupId || `group-${nameEntities[longest].id}`;
        targetP = nameEntities[longest].pseudonym;
      }

      for (const idx of component) {
        const e = nameEntities[idx];
        e.groupId = targetG;
        e.pseudonym = targetP || e.pseudonym;
        e.treated = isTreated || e.treated;
      }
    }
  }

  return newEntities;
}

export function anonymizeText(text: string, entities: PIIEntity[]): string {
  let result = text;
  
  // Sort entities by length descending to avoid partial replacements
  // Include treated entities even if not "enabled" by the auto-scanner
  const sortedEntities = [...entities]
    .filter(e => (e.enabled || e.treated) && !e.ignored)
    .sort((a, b) => b.original.length - a.original.length);

  // Use a map to avoid duplicate replacements of the same original text
  const handled = new Set<string>();

  sortedEntities.forEach(entity => {
    const original = entity.original.trim();
    if (!original || handled.has(original.toLowerCase())) return;
    handled.add(original.toLowerCase());

    // Escape special characters and allow flexible spacing (handles justified text and extra spaces)
    const escaped = original
      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      .replace(/\s+/g, '\\s+');
    
    // Use word boundaries that respect Unicode characters (Portuguese accents)
    // The standard \b does not handle 'é', 'à', etc. correctly as word characters.
    const wordChars = 'a-zA-ZÀ-ÿ0-9_';
    const regex = new RegExp(`(?<![${wordChars}])${escaped}(?![${wordChars}])`, 'gi');
    result = result.replace(regex, entity.pseudonym);
  });

  return result;
}
