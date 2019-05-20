module Viewpoint::EWS::Types
  class CalendarFolder
    include Viewpoint::EWS
    include Viewpoint::EWS::Types
    include Viewpoint::EWS::Types::GenericFolder

    # Fetch items between a given time period
    # @param [DateTime] start_date the time to start fetching Items from
    # @param [DateTime] end_date the time to stop fetching Items from
    # @param [Boolean] calendar_view indicates if the calendar view element should be passed see https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/recurrence-patterns-and-ews#expanded-vs-non-expanded-views
    def items_between(start_date, end_date, calendar_view = false, opts={})
      if calendar_view
        opts.merge!({ calendar_view: { start_date: start_date, end_date: end_date } })
        items(opts)
      else
        items(opts) do |obj|
          obj.restriction = { :and =>
            [
              {:is_greater_than_or_equal_to =>
                [
                  {:field_uRI => {:field_uRI=>'calendar:Start'}},
                  {:field_uRI_or_constant=>{:constant => {:value =>start_date}}}
                ]
              },
              {:is_less_than_or_equal_to =>
                [
                  {:field_uRI => {:field_uRI=>'calendar:End'}},
                  {:field_uRI_or_constant=>{:constant => {:value =>end_date}}}
                ]
              }
            ]
          }
        end
      end
    end

    # Creates a new appointment
    # @param attributes [Hash] Parameters of the calendar item. Some example attributes are listed below.
    # @option attributes :subject [String]
    # @option attributes :start [Time]
    # @option attributes :end [Time]
    # @return [CalendarItem]
    # @see Template::CalendarItem
    def create_item(attributes, to_ews_create_opts = {})
      template = Viewpoint::EWS::Template::CalendarItem.new attributes
      template.saved_item_folder_id = {id: self.id, change_key: self.change_key}
      rm = ews.create_item(template.to_ews_create(to_ews_create_opts)).response_messages.first
      if rm && rm.success?
        CalendarItem.new ews, rm.items.first[:calendar_item][:elems].first
      else
        raise EwsCreateItemError, "Could not create item in folder. #{rm.code}: #{rm.message_text}" unless rm
      end
    end

  end
end
