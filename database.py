from sqlalchemy import Column, String, UUID
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from sqlalchemy import create_engine
from uuid import uuid4 as uuid
from sqlalchemy import and_, delete

Base = declarative_base()
class Media(Base):
    __tablename__ = 'media'
    id = Column(UUID, nullable=False, primary_key=True, comment='ID Созданного медиа')
    author_id = Column(UUID, nullable=False, comment='ID Приложения, создавшего медиа')
    name = Column(String, nullable=False, comment='Название файла')
    url = Column(String, comment='Прямая ссылка на медиа')
    author_ver = Column(String, comment='Версия приложения, создавшего медиа')
    resource_id = Column(String, nullable=False, comment='Идентификатор, по которому можно найти медиа')
    product_id = Column(String, nullable=False, comment='Идентификатор товара')
    description = Column(String, comment='Дополнительная информация о медиа')

class DBConnect:
    def __init__(self, appInfo: dict):
        self.login = appInfo.get('DBLogin')
        self.password = appInfo.get('DBPassword')
        self.appVer = appInfo.get('AppVer')
        self.id = appInfo.get('DBID')
        try:
            engine = create_engine(url=f"postgresql+psycopg2://{self.login}:{self.password}@scope-db-lego-invest-group.db-msk0.amvera.tech:5432/LEGOSystems")
            Session = sessionmaker(bind=engine)
            self.session = Session()
        except Exception as e:
            raise SystemError(f'Ошибка авторизации в базе данных: {e}')

    def create_media(self, url: str, filename: str, resource_id: str, product_id: str, description: str | None):
        new_media = Media(
            id = uuid(),
            author_id = self.id,
            author_ver = self.appVer,
            resource_id = resource_id,
            product_id = product_id,
            url = url,
            name = filename,
            description = description
        )
        self.session.add(new_media)
        self.session.commit()

    def is_actual_media_generated(self, resource_id: str):
        results = self.session.query(Media.author_ver).where(and_(Media.author_id == self.id, Media.resource_id == resource_id)).all()
        if len(results) == 0: return False
        return all(res[0] == self.appVer for res in results)

    def delete_media(self, resource_id: str, filename: str):
        media = self.session.query(Media).where(and_(Media.author_id == self.id, Media.resource_id == resource_id, Media.name == filename)).one_or_none()
        if media is not None:
            self.session.execute(delete(Media).where(Media.id == media.id))
    
    def close(self):
        self.session.close()