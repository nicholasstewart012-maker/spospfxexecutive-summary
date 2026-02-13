import * as React from 'react';
import styles from './ExecutiveSummary.module.scss';
import { IExecutiveSummaryProps } from './IExecutiveSummaryProps';
import { SPService } from '../../../services/SPService';
import { IEvent } from '../../../common/IEvent';
import { getCategoryColor, AppConstants } from '../../../common/Constants';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import moment from 'moment-timezone';
import * as DOMPurify from 'dompurify';

const ExecutiveSummary: React.FC<IExecutiveSummaryProps> = (props) => {
  const [events, setEvents] = React.useState<IEvent[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>("");

  React.useEffect(() => {
    const fetchData = async () => {
      if (!props.listName || !props.siteUrl) {
        setError("Please configure the Site URL and List Name in the Web Part properties.");
        setLoading(false);
        return;
      }

      try {
        setLoading(true);
        const spService = new SPService(props.context);
        const fetchedEvents = await spService.getEvents(props.siteUrl, props.listName);
        setEvents(fetchedEvents);
        setError("");
      } catch (err: any) {
        setError("Failed to load events. Please check the console for details.");
        console.error(err);
      } finally {
        setLoading(false);
      }
    };

    void fetchData();
  }, [props.listName, props.siteUrl, props.context]);

  const groupedEvents = React.useMemo(() => {
    const groups: { [key: string]: IEvent[] } = {};
    events.forEach(event => {
      const category = event.Category || AppConstants.Categories.Other;
      if (!groups[category]) {
        groups[category] = [];
      }
      groups[category].push(event);
    });
    return groups;
  }, [events]);

  const renderEventDate = (dateStr: string) => {
    const date = moment(dateStr);
    return (
      <div className={styles.dateBox}>
        <span className={styles.month}>{date.format("MMM")}</span>
        <span className={styles.day}>{date.format("DD")}</span>
      </div>
    );
  };

  const categories = Object.keys(groupedEvents);

  // Helper to get total counts per category
  const getCategoryCount = (category: string) => {
    return groupedEvents[category]?.length || 0;
  }

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Loading events..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  return (
    <section className={styles.executiveSummary}>
      <div className={styles.header}>
        <h2>Executive Summary</h2>
        <p>Overview of upcoming events and training.</p>
      </div>

      <div className={styles.dashboard}>
        {/* Summary Cards */}
        {Object.keys(AppConstants.Categories).map((key) => {
          const cat = AppConstants.Categories[key as keyof typeof AppConstants.Categories];
          return (
            <div key={cat} className={styles.summaryCard} style={{ borderTopColor: getCategoryColor(cat) }}>
              <div className={styles.cardCount} style={{ color: getCategoryColor(cat) }}>{getCategoryCount(cat)}</div>
              <div className={styles.cardLabel}>{cat}</div>
            </div>
          );
        })}
      </div>

      <div className={styles.eventsContainer}>
        {categories.length === 0 && <div>No upcoming events found.</div>}
        {categories.map(category => (
          <div key={category} className={styles.categorySection}>
            <h3 className={styles.categoryHeader} style={{ color: getCategoryColor(category), borderBottomColor: getCategoryColor(category) }}>
              {category}
            </h3>
            <div className={styles.eventGrid}>
              {groupedEvents[category].map(event => (
                <div key={event.Id} className={styles.eventCard}>
                  <div className={styles.eventDate}>
                    {renderEventDate(event.EventDate)}
                  </div>
                  <div className={styles.eventDetails}>
                    <h4 className={styles.eventTitle}>{event.Title}</h4>

                    <div className={styles.timezones}>
                      {event.StartTimeZoneCST && <span><strong>CST:</strong> {event.StartTimeZoneCST} - {event.EndTimeZoneCST}</span>}
                      {event.StartTimeZoneEST && <span><strong>EST:</strong> {event.StartTimeZoneEST} - {event.EndTimeZoneEST}</span>}
                    </div>

                    {event.Location && <div className={styles.metaLine}><strong>Location:</strong> {event.Location}</div>}
                    {event.TargetAudience && <div className={styles.metaLine}><strong>Audience:</strong> {event.TargetAudience}</div>}

                    <div className={styles.description} dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(event.Description) }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </section>
  );
};

export default ExecutiveSummary;
